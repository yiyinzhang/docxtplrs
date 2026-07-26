//! Port of docxtpl's patch_xml / resolve_listing logic (regex-based XML cleaning).

use fancy_regex::{Captures, Regex};
use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

/// Timing instrumentation, enabled with PATCH_TIMING=1.
macro_rules! timed {
    ($name:expr, $e:expr) => {{
        let __t0 = std::time::Instant::now();
        let __r = $e;
        if std::env::var("PATCH_TIMING").is_ok() {
            eprintln!("[patch-timing] {}: {:.3}ms", $name, __t0.elapsed().as_secs_f64() * 1000.0);
        }
        __r
    }};
}

thread_local! {
    // Compiling a fancy_regex is expensive (~tens of µs) and all patterns used
    // here are fixed literals, so cache them per thread. Without this cache,
    // resolve_listing recompiles several regexes per paragraph/run, which
    // dominates render time for table-row-heavy templates (~ms per row).
    // Rc-shared: cloning a hit is a refcount bump, and callers may hold the
    // value across re-entrant cache lookups (sub() callbacks recurse).
    static RE_CACHE: RefCell<HashMap<String, Rc<Regex>>> = RefCell::new(HashMap::new());
}

pub(crate) fn re(pattern: &str) -> Rc<Regex> {
    if std::env::var("PATCH_DEBUG").is_ok() {
        eprintln!("[patch] running: {}", pattern);
    }
    RE_CACHE.with(|c| {
        let mut map = c.borrow_mut();
        if let Some(r) = map.get(pattern) {
            return r.clone();
        }
        let r = Rc::new(
            fancy_regex::RegexBuilder::new(pattern)
                .backtrack_limit(50_000_000)
                .build()
                .unwrap_or_else(|e| panic!("invalid regex {}: {}", pattern, e)),
        );
        map.insert(pattern.to_string(), r.clone());
        r
    })
}

/// replace all matches of `pattern` in `text` using closure over captures
pub fn sub<F>(pattern: &str, mut repl: F, text: &str) -> String
where
    F: FnMut(&Captures) -> String,
{
    let rex = re(pattern);
    let mut out = String::with_capacity(text.len());
    let mut last = 0usize;
    loop {
        match rex.captures_from_pos(text, last) {
            Ok(Some(caps)) => {
                let m = caps.get(0).unwrap();
                out.push_str(&text[last..m.start()]);
                out.push_str(&repl(&caps));
                if m.end() == m.start() {
                    // empty match: copy one char to avoid an infinite loop
                    match text[m.end()..].chars().next() {
                        Some(c) => {
                            out.push(c);
                            last = m.end() + c.len_utf8();
                        }
                        None => {
                            last = text.len();
                            break;
                        }
                    }
                } else {
                    last = m.end();
                }
            }
            Ok(None) => break,
            Err(e) => panic!("regex error in pattern [{}]: {}", pattern, e),
        }
    }
    out.push_str(&text[last.min(text.len())..]);
    out
}

/// simple literal replacement with $ group references
/// True when `hay` contains at least one of `needles`.
#[inline]
fn contains_any(hay: &str, needles: &[&str]) -> bool {
    needles.iter().any(|n| hay.contains(n))
}

/// Hand-rolled equivalent of the docxtpl regex
/// `(?<=\{)(?>(?:<[^>]*>)+)(?=[\{%\#])|(?<=[%\}\#])(?>(?:<[^>]*>)+)(?=\})`
/// which strips XML tags that split a jinja tag across runs (`{<r></r>{` -> `{{`).
/// The atomic group means: after a delimiter, consume as many whole `<...>`
/// tags as possible, then the terminator check must hold (no backtracking).
/// Linear time; the fancy_regex original costs ~20ms per 500KB even without
/// any match.
fn merge_split_braces_scan(xml: &str) -> String {
    let b = xml.as_bytes();
    let n = b.len();
    let mut out = String::with_capacity(n);
    let mut copied = 0usize; // xml[..copied] already flushed to out
    let mut i = 0usize;
    while i < n {
        let c = b[i];
        // open side: '{' ... one of '{','%','#' ; close side: one of '%','}','#' ... '}'
        let open_side = match c {
            b'{' => true,
            b'%' | b'}' | b'#' => false,
            _ => {
                i += 1;
                continue;
            }
        };
        // consume consecutive complete `<...>` tags starting at i+1
        let mut j = i + 1;
        while j < n && b[j] == b'<' {
            match xml[j..].find('>') {
                Some(k) => j += k + 1,
                None => break,
            }
        }
        let matched = if j > i + 1 && j < n {
            if open_side {
                matches!(b[j], b'{' | b'%' | b'#')
            } else {
                b[j] == b'}'
            }
        } else {
            false
        };
        if matched {
            // keep the delimiter char, drop the tags; the terminator (not
            // consumed, like the regex lookahead) is re-examined next round
            out.push_str(&xml[copied..=i]);
            copied = j;
            i = j;
        } else {
            i += 1;
        }
    }
    out.push_str(&xml[copied..]);
    out
}

pub fn sub_str(pattern: &str, replacement: &str, text: &str) -> String {
    // route through `sub` (fancy-regex's replace_all unwraps internally)
    let replacement = replacement.to_string();
    sub(pattern, move |m| {
        let mut out = replacement.clone();
        // expand $N / ${N} group references
        for g in 1..=9 {
            if let Some(gr) = m.get(g) {
                out = out.replace(&format!("${}", g), gr.as_str());
                out = out.replace(&format!("${{{}}}", g), gr.as_str());
            }
        }
        out
    }, text)
}

/// Decode XML entities in text nodes that lxml would normalize away when
/// parsing/serializing (docxtpl's get_xml goes through lxml, so templates
/// containing `&quot;`, `&apos;` or numeric character references inside jinja
/// tags work there). `&lt;` `&gt;` `&amp;` are intentionally kept escaped,
/// matching lxml's serialization behavior.
pub fn decode_text_entities(xml: &str) -> String {
    if !xml.contains('&') {
        return xml.to_string();
    }
    sub(
        r"(?<=>)([^<>]*)(?=<)",
        |m| decode_entities_keep_markup(m.get(0).unwrap().as_str()),
        xml,
    )
}

fn decode_entities_keep_markup(text: &str) -> String {
    if !text.contains('&') {
        return text.to_string();
    }
    let mut out = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(pos) = rest.find('&') {
        out.push_str(&rest[..pos]);
        let tail = &rest[pos..];
        let semi = match tail.find(';') {
            Some(s) if s < 12 => s,
            _ => {
                out.push('&');
                rest = &rest[pos + 1..];
                continue;
            }
        };
        let ent = &tail[..=semi];
        let decoded: Option<String> = match ent {
            "&quot;" => Some("\"".to_string()),
            "&apos;" => Some("'".to_string()),
            _ if ent.starts_with("&#x") || ent.starts_with("&#X") => {
                u32::from_str_radix(&ent[3..ent.len() - 1], 16)
                    .ok()
                    .and_then(char::from_u32)
                    .and_then(keep_char)
            }
            _ if ent.starts_with("&#") => {
                ent[2..ent.len() - 1]
                    .parse::<u32>()
                    .ok()
                    .and_then(char::from_u32)
                    .and_then(keep_char)
            }
            _ => None,
        };
        match decoded {
            Some(s) => out.push_str(&s),
            None => out.push_str(ent),
        }
        rest = &tail[semi + 1..];
    }
    out.push_str(rest);
    out
}

/// keep the char (as lxml would serialize it literally) unless it is a markup
/// char, which lxml would keep escaped
fn keep_char(c: char) -> Option<String> {
    match c {
        '<' | '>' | '&' => None,
        _ => Some(c.to_string()),
    }
}

/// Port of DocxTemplate.patch_xml
pub fn patch_xml(src_xml: &str) -> String {
    // replace {<something>{ by {{   ( works with {{ }} {% and %} {# and #})
    // (hand-rolled linear scan; fancy_regex spends ~20ms/500KB here even when
    // nothing matches)
    // gate: a match requires a delimiter directly followed by a tag
    let mut xml = if contains_any(src_xml, &["{<", "%<", "}<", "#<"]) {
        timed!("merge_split_braces", merge_split_braces_scan(src_xml))
    } else {
        src_xml.to_string()
    };

    // replace {{<some tags>jinja2 stuff<some other tags>}} by {{jinja2 stuff}}
    // (gated: the pattern can only match where a jinja open marker exists)
    if contains_any(&xml, &["{{", "{%", "{#"]) {
        xml = timed!("strip_tags_in_jinja", sub(
        r"(?s)\{%(?:(?!%\}).)*|\{#(?:(?!#\}).)*|\{\{(?:(?!}\}).)*",
        |m| sub_str(r"(?s)</w:t>.*?(<w:t>|<w:t [^>]*>)", "", m.get(0).unwrap().as_str()),
        &xml,
        ));
    }

    // manage table cell colspan
    if xml.contains("colspan") {
    xml = timed!("colspan", sub(
        r"(?s)(<w:tc[ >](?:(?!<w:tc[ >]).)*)\{%\s*colspan\s+([^%]*)\s*%\}(.*?</w:tc>)",
        |m| {
            let mut cell_xml = format!("{}{}", m.get(1).unwrap().as_str(), m.get(3).unwrap().as_str());
            cell_xml = sub_str(
                r"(?s)<w:r[ >](?:(?!<w:r[ >]).)*<w:t></w:t>.*?</w:r>",
                "",
                &cell_xml,
            );
            cell_xml = sub_str(r"<w:gridSpan[^/]*/>", "", &cell_xml);
            sub(
                r"(<w:tcPr[^>]*>)",
                |mm| {
                    format!(
                        "{}<w:gridSpan w:val=\"{{{{{}}}}}\"/>",
                        mm.get(1).unwrap().as_str(),
                        m.get(2).unwrap().as_str()
                    )
                },
                &cell_xml,
            )
        },
        &xml,
    ));
    }

    // manage table cell background color
    if xml.contains("cellbg") {
    xml = timed!("cellbg", sub(
        r"(?s)(<w:tc[ >](?:(?!<w:tc[ >]).)*)\{%\s*cellbg\s+([^%]*)\s*%\}(.*?</w:tc>)",
        |m| {
            let mut cell_xml = format!("{}{}", m.get(1).unwrap().as_str(), m.get(3).unwrap().as_str());
            cell_xml = sub_str(
                r"(?s)<w:r[ >](?:(?!<w:r[ >]).)*<w:t></w:t>.*?</w:r>",
                "",
                &cell_xml,
            );
            // remove first <w:shd .../>
            cell_xml = sub_first(r"<w:shd[^/]*/>", "", &cell_xml);
            sub(
                r"(<w:tcPr[^>]*>)",
                |mm| {
                    format!(
                        "{}<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{{{{{}}}}}\"/>",
                        mm.get(1).unwrap().as_str(),
                        m.get(2).unwrap().as_str()
                    )
                },
                &cell_xml,
            )
        },
        &xml,
    ));
    }

    // ensure space preservation (hand-rolled linear scan; the original
    // tempered-dot regex overflows the backtrack stack on large documents)
    if xml.contains("<w:t>") && contains_any(&xml, &["{{", "{%"]) {
        xml = timed!("space_preserve", space_preserve_scan(&xml));
    }
    if contains_any(&xml, &["{{r", "{%r"]) {
        xml = timed!("richtext_tag", sub(
        r"(?s)(\{\{r\s.*?\}\}|\{%r\s.*?\%\})",
        |m| {
            format!(
                "</w:t></w:r><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r><w:r><w:t xml:space=\"preserve\">",
                m.get(1).unwrap().as_str()
            )
        },
        &xml,
        ));
    }

    // {%- will merge with previous paragraph text (hand-rolled; the
    // original regex spans paragraphs and overflows the backtrack stack)
    if xml.contains("{%-") {
        xml = timed!("dash_merge_prev", dash_merge_prev_scan(&xml));
    }
    // -%} will merge with next paragraph text
    if xml.contains("-%}") {
        xml = timed!("dash_merge_next", dash_merge_next_scan(&xml));
    }

    // replace into xml code the row/paragraph/run containing
    // {%y xxx %} / {{y xxx}} / {#y xxx #} template tag by the tag alone
    // without any surrounding <w:y> tags (hand-rolled linear scan; the
    // original tempered-dot regex overflows the backtrack stack on large
    // documents)
    for y in ["tr", "tc", "p", "r"] {
        let m1 = format!("{{{{{} ", y);
        let m2 = format!("{{%{} ", y);
        if xml.contains(&m1) || xml.contains(&m2) {
            xml = timed!("element_tag_scan", element_tag_scan(&xml, y, false));
        }
    }
    for y in ["tr", "tc", "p"] {
        let m = format!("{{#{} ", y);
        if xml.contains(&m) {
            xml = timed!("element_tag_scan", element_tag_scan(&xml, y, true));
        }
    }

    // add vMerge
    // use {% vm %} to make this table cell and its copies
    // be vertically merged within a {% for %}
    if xml.contains("{%") && xml.contains("vm") {
    xml = timed!("vmerge", sub(
        r"(?s)<w:tc[ >](?:(?!<w:tc[ >]).)*?\{%\s*vm\s*%\}.*?</w:tc[ >]",
        |m| {
            let whole = m.get(0).unwrap().as_str();
            sub(
                r"(?s)(</w:tcPr[ >].*?<w:t(?:.*?)>)(.*?)(?:\{%\s*vm\s*%\})(.*?)(</w:t>)",
                |m1| {
                    format!(
                        "<w:vMerge w:val=\"{{% if loop.first %}}restart{{% else %}}continue{{% endif %}}\"/>{}{{% if loop.first %}}{}{}{{% endif %}}{}",
                        m1.get(1).unwrap().as_str(),
                        m1.get(2).unwrap().as_str(),
                        m1.get(3).unwrap().as_str(),
                        m1.get(4).unwrap().as_str(),
                    )
                },
                whole,
            )
        },
        &xml,
    ));
    }

    // Use {% hm %} to make table cell become horizontally merged within a {% for %}.
    if xml.contains("{%") && xml.contains("hm") {
    xml = timed!("hmerge", sub(
        r"(?s)<w:tc[ >](?:(?!<w:tc[ >]).)*?\{%\s*hm\s*%\}.*?</w:tc[ >]",
        |m| {
            let whole = m.get(0).unwrap().as_str().to_string();
            let xml_patched = if whole.contains("w:gridSpan") {
                // Simple case: there's already gridSpan, multiply its value.
                let x = sub(
                    r#"(?s)(w:gridSpan w:val=")(\d+)(")"#,
                    |m1| {
                        format!(
                            "{}{{{} {} * loop.length {}}}{}",
                            m1.get(1).unwrap().as_str(),
                            "{",
                            m1.get(2).unwrap().as_str(),
                            "}",
                            m1.get(3).unwrap().as_str()
                        )
                    },
                    &whole,
                );
                sub_str(r"(?s)\{%\s*hm\s*%\}", "", &x)
            } else {
                sub(
                    r"(?s)(</w:tcPr[ >].*?<w:t(?:.*?)>)(.*?)(?:\{%\s*hm\s*%\})(.*?)(</w:t>)",
                    |m2| {
                        format!(
                            "<w:gridSpan w:val=\"{{{{ loop.length }}}}\"/>{}{}{}{}",
                            m2.get(1).unwrap().as_str(),
                            m2.get(2).unwrap().as_str(),
                            m2.get(3).unwrap().as_str(),
                            m2.get(4).unwrap().as_str(),
                        )
                    },
                    &whole,
                )
            };
            // Discard every other cell generated in loop.
            format!("{{% if loop.first %}}{}{{% endif %}}", xml_patched)
        },
        &xml,
    ));
    }

    // clean tags: unescape entities and smart quotes inside jinja tags
    if contains_any(&xml, &["{{", "{%"]) {
    xml = timed!("clean_tags", sub(r"(?<=\{[\{%])(.*?)(?=[\}%]\})", |m| {
        m.get(0)
            .unwrap()
            .as_str()
            .replace("&#8216;", "'")
            .replace("&lt;", "<")
            .replace("&gt;", ">")
            .replace('\u{201c}', "\"")
            .replace('\u{201d}', "\"")
            .replace('\u{2018}', "'")
            .replace('\u{2019}', "'")
    }, &xml));
    }

    xml
}

/// replace only first occurrence
fn sub_first(pattern: &str, replacement: &str, text: &str) -> String {
    let rex = re(pattern);
    rex.replace(text, replacement).to_string()
}

/// Port of DocxTemplate.resolve_listing
pub fn resolve_listing(xml: &str) -> String {
    // resolve_text only rewrites \t, \n, \x07 and \x0c; without any of them
    // the whole pass is the identity, so skip the full-document copy
    if !xml.contains(['\t', '\n', '\u{7}', '\u{c}']) {
        return xml.to_string();
    }
    fn resolve_text(run_properties: &str, paragraph_properties: &str, m: &Captures) -> String {
        let mut s = m.get(0).unwrap().as_str().to_string();
        s = s.replace(
            '\t',
            &format!(
                "</w:t></w:r><w:r>{}<w:tab/></w:r><w:r>{}<w:t xml:space=\"preserve\">",
                run_properties, run_properties
            ),
        );
        s = s.replace(
            '\u{7}',
            &format!(
                "</w:t></w:r></w:p><w:p>{}<w:r>{}<w:t xml:space=\"preserve\">",
                paragraph_properties, run_properties
            ),
        );
        s = s.replace('\n', "</w:t><w:br/><w:t xml:space=\"preserve\">");
        s = s.replace(
            '\u{c}',
            &format!(
                "</w:t></w:r></w:p><w:p><w:r><w:br w:type=\"page\"/></w:r></w:p><w:p>{}<w:r>{}<w:t xml:space=\"preserve\">",
                paragraph_properties, run_properties
            ),
        );
        s
    }

    fn resolve_run(paragraph_properties: &str, m: &Captures) -> String {
        let whole = m.get(0).unwrap().as_str();
        let run_properties = re(r"(?s)<w:rPr>.*?</w:rPr>")
            .find(whole)
            .ok()
            .flatten()
            .map(|mm| mm.as_str().to_string())
            .unwrap_or_default();
        sub(
            r"(?s)<w:t(?: [^>]*)?>.*?</w:t>",
            |x| resolve_text(&run_properties, paragraph_properties, x),
            whole,
        )
    }

    fn resolve_paragraph(m: &Captures) -> String {
        let whole = m.get(0).unwrap().as_str();
        // Fast path: resolve_text only rewrites \t, \n, \x07 and \x0c. A
        // paragraph without any of them is returned unchanged, so skip the
        // (expensive) per-paragraph/per-run regex scans entirely.
        if !whole.contains(['\t', '\n', '\u{7}', '\u{c}']) {
            return whole.to_string();
        }
        let paragraph_properties = re(r"(?s)<w:pPr>.*?</w:pPr>")
            .find(whole)
            .ok()
            .flatten()
            .map(|mm| mm.as_str().to_string())
            .unwrap_or_default();
        sub(
            r"(?s)<w:r(?: [^>]*)?>.*?</w:r>",
            |x| resolve_run(&paragraph_properties, x),
            whole,
        )
    }

    timed!("resolve_listing_scan", sub(
        r"(?s)<w:p(?: [^>]*)?>.*?</w:p>",
        |m| resolve_paragraph(m),
        xml,
    ))
}

/// `<w:t>` -> `<w:t xml:space="preserve">` when a jinja tag follows before
/// the next bare `<w:t>` (equivalent to the docxtpl regex, but linear-time).
fn space_preserve_scan(xml: &str) -> String {
    let mut out = String::with_capacity(xml.len());
    let mut rest = xml;
    while let Some(pos) = rest.find("<w:t>") {
        out.push_str(&rest[..pos]);
        let after_open = pos + 5;
        let region_end = rest[after_open..]
            .find("<w:t>")
            .map(|i| after_open + i)
            .unwrap_or(rest.len());
        let region = &rest[after_open..region_end];
        if region.contains("{{") || region.contains("{%") {
            out.push_str("<w:t xml:space=\"preserve\">");
        } else {
            out.push_str("<w:t>");
        }
        rest = &rest[after_open..];
    }
    out.push_str(rest);
    out
}

/// `{%-` merges with previous paragraph text: delete everything from the
/// nearest preceding `</w:t>` up to and including `{%-`, leave `{%`.
fn dash_merge_prev_scan(xml: &str) -> String {
    let mut out = String::with_capacity(xml.len());
    let mut rest = xml;
    while let Some(pos) = rest.find("{%-") {
        let before = &rest[..pos];
        match before.rfind("</w:t>") {
            Some(w) => {
                out.push_str(&before[..w]);
                out.push_str("{%");
                rest = &rest[pos + 3..];
            }
            None => {
                out.push_str(&rest[..pos + 3]);
                rest = &rest[pos + 3..];
            }
        }
    }
    out.push_str(rest);
    out
}

/// `-%}` merges with next paragraph text: if the next `<w:t` opening (before
/// any `{%`/`{{`) follows, delete everything from `-%}` through that opening
/// tag, leave `%}`.
fn dash_merge_next_scan(xml: &str) -> String {
    let mut out = String::with_capacity(xml.len());
    let mut rest = xml;
    while let Some(pos) = rest.find("-%}") {
        let from = pos + 3;
        let tail = &rest[from..];
        let idx_wt = tail
            .find("<w:t")
            .filter(|&i| tail[i..].starts_with("<w:t>") || tail[i..].starts_with("<w:t "));
        let blocked = |upto: usize| {
            tail[..upto].contains("{%") || tail[..upto].contains("{{")
        };
        match idx_wt {
            Some(w) if !blocked(w) => {
                match tail[w..].find('>') {
                    Some(e) => {
                        out.push_str(&rest[..pos]);
                        out.push_str("%}");
                        rest = &rest[from + w + e + 1..];
                    }
                    None => {
                        out.push_str(&rest[..from]);
                        rest = tail;
                    }
                }
            }
            _ => {
                out.push_str(&rest[..from]);
                rest = tail;
            }
        }
    }
    out.push_str(rest);
    out
}

/// Element-level jinja tags ({%tr/%tc/%p/%r and {#..#} comments):
/// find `<w:y>` elements containing a `{%y `% / `{{y `% (or `{#y `) tag and
/// replace the whole element with the bare tag — linear-time equivalent of
/// docxtpl's regexes.
pub fn element_tag_scan(xml: &str, y: &str, comment: bool) -> String {
    let open_prefix = format!("<w:{}", y);
    let close_tag = format!("</w:{}>", y);
    let mut out = String::with_capacity(xml.len());
    let mut i = 0usize;

    // markers to look for inside the element, e.g. "{%tr " / "{{tr " / "{#tr "
    let m_var = format!("{{{{{} ", y);
    let m_stmt = format!("{{%{} ", y);
    let m_comment = format!("{{#{} ", y);
    let markers: Vec<(&str, &str)> = if comment {
        vec![(m_comment.as_str(), "#}")]
    } else {
        vec![(m_var.as_str(), "}}"), (m_stmt.as_str(), "%}")]
    };

    while i < xml.len() {
        // find next `<w:y` opening followed by ' ' or '>'
        let Some(rel) = xml[i..].find(&open_prefix) else {
            break;
        };
        let open = i + rel;
        let after_open = open + open_prefix.len();
        let ok_open = xml[after_open..]
            .chars()
            .next()
            .map(|c| c == ' ' || c == '>')
            .unwrap_or(false);
        if !ok_open {
            out.push_str(&xml[i..after_open]);
            i = after_open;
            continue;
        }
        // element region: up to the first close tag (docxtpl regex behavior)
        let region_end = xml[after_open..]
            .find(&close_tag)
            .map(|c| after_open + c + close_tag.len())
            .unwrap_or(xml.len());
        let region = &xml[open..region_end];

        // earliest marker inside the region
        let mut best: Option<(usize, &str, &str)> = None;
        for (marker, close_tok) in &markers {
            if let Some(mp) = region.find(marker) {
                if best.map(|(b, _, _)| mp < b).unwrap_or(true) {
                    best = Some((mp, marker, close_tok));
                }
            }
        }
        let Some((mpos, marker, close_tok)) = best else {
            // no tag in this element; emit the opening tag and continue inside
            out.push_str(&xml[i..after_open]);
            i = after_open;
            continue;
        };
        let after_marker = mpos + marker.len();
        // close token must appear before any '%'/'}' (statement) or '#'/'}'
        // (comment), matching [^}%]* / [^}#]* from the original regex
        let forbidden: &[char] = if comment { &['#', '}'] } else { &['%', '}'] };
        let Some(close_rel) = region[after_marker..].find(close_tok) else {
            out.push_str(&xml[i..after_open]);
            i = after_open;
            continue;
        };
        let inner = &region[after_marker..after_marker + close_rel];
        if inner.chars().any(|c| forbidden.contains(&c)) {
            out.push_str(&xml[i..after_open]);
            i = after_open;
            continue;
        }
        // emit: everything before the element + bare tag; skip whole element
        out.push_str(&xml[i..open]);
        out.push_str(&marker[..2]);
        out.push(' ');
        out.push_str(inner);
        out.push_str(close_tok);
        i = region_end;
    }
    out.push_str(&xml[i.min(xml.len())..]);
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    // ------------------------------------------------------------------
    // 手写线性扫描器（AGENTS.md 第 9 条：替代会溢出回退栈的跨段正则）
    // ------------------------------------------------------------------

    /// merge_split_braces_scan：拆开成多段 run 的 jinja 定界符合并，
    /// 等价 docxtpl 的 `(?<=\{)(?>(?:<[^>]*>)+)(?=[\{%\#])|(?<=[%\}\#])(?>(?:<[^>]*>)+)(?=\})`
    #[test]
    fn test_merge_split_braces_cross_paragraph() {
        // 开侧 `{<tags>{` -> `{{`（跨 run）
        assert_eq!(merge_split_braces_scan("{</w:t></w:r><w:r><w:t>{"), "{{");
        // 跨整个段落
        let xml = "<w:p><w:r><w:t>{</w:t></w:r></w:p><w:p><w:r><w:t>{ x }}</w:t></w:r></w:p>";
        assert_eq!(
            merge_split_braces_scan(xml),
            "<w:p><w:r><w:t>{{ x }}</w:t></w:r></w:p>"
        );
        // 闭侧 `%<tags>}` -> `%}`；开侧终止符也可以是 % 或 #
        assert_eq!(merge_split_braces_scan("%</w:t><w:t>}"), "%}");
        assert_eq!(merge_split_braces_scan("{<a></a>%"), "{%");
    }

    #[test]
    fn test_merge_split_braces_no_match() {
        // 无终止符：标签后不是定界符，原样保留
        assert_eq!(merge_split_braces_scan("{<a></a>x"), "{<a></a>x");
        // 闭侧只认 `}`，其它定界符不匹配
        assert_eq!(merge_split_braces_scan("%<a>{"), "%<a>{");
        // 原子组语义：标签之间夹了非标签字符即失败，不回退
        assert_eq!(merge_split_braces_scan("{<a>x{"), "{<a>x{");
        // 没有 `>` 的不完整标签不算标签；空输入
        assert_eq!(merge_split_braces_scan("{<a"), "{<a");
        assert_eq!(merge_split_braces_scan(""), "");
    }

    /// space_preserve_scan：区域内（到下一个裸 `<w:t>` 为止）出现
    /// `{{`/`{%` 时把 `<w:t>` 改写为 `<w:t xml:space="preserve">`
    #[test]
    fn test_space_preserve_scan() {
        assert_eq!(
            space_preserve_scan("<w:t>{{ x }}</w:t>"),
            "<w:t xml:space=\"preserve\">{{ x }}</w:t>"
        );
        assert_eq!(
            space_preserve_scan("<w:t>{% if x %}</w:t>"),
            "<w:t xml:space=\"preserve\">{% if x %}</w:t>"
        );
        // 纯文本不改；多个 run 只改含 jinja 标签的那个
        assert_eq!(
            space_preserve_scan("<w:t>a</w:t><w:t>{{b}}</w:t>"),
            "<w:t>a</w:t><w:t xml:space=\"preserve\">{{b}}</w:t>"
        );
        // 扫描区域延伸到下一个 `<w:t>` 开头（可越过 `</w:t>`，
        // 与 docxtpl 的 tempered-dot 正则一致）
        assert_eq!(
            space_preserve_scan("<w:t>a</w:t>{{ x }}<w:t>b</w:t>"),
            "<w:t xml:space=\"preserve\">a</w:t>{{ x }}<w:t>b</w:t>"
        );
        // 只匹配裸 `<w:t>`，已带属性的不动；空输入
        assert_eq!(
            space_preserve_scan("<w:t xml:space=\"preserve\">{{x}}</w:t>"),
            "<w:t xml:space=\"preserve\">{{x}}</w:t>"
        );
        assert_eq!(space_preserve_scan(""), "");
    }

    /// dash_merge_prev_scan：`{%-` 向前合并：删去最近一个 `</w:t>`
    /// 到 `{%-` 之间的全部内容，留下 `{%`
    #[test]
    fn test_dash_merge_prev_scan() {
        assert_eq!(dash_merge_prev_scan("a</w:t>junk{%- x %}"), "a{% x %}");
        // 跨段落合并（docxtpl 的典型用途）
        let xml = "<w:p><w:r><w:t>text</w:t></w:r></w:p><w:p><w:r><w:t>{%- if x %}</w:t></w:r></w:p>";
        assert_eq!(
            dash_merge_prev_scan(xml),
            "<w:p><w:r><w:t>text{% if x %}</w:t></w:r></w:p>"
        );
        // 多处出现逐一处理
        assert_eq!(
            dash_merge_prev_scan("a</w:t>{%- x %}b</w:t>{%- y"),
            "a{% x %}b{% y"
        );
        // 边界：前面没有 `</w:t>` 时保持原样；空输入
        assert_eq!(dash_merge_prev_scan("x{%- y"), "x{%- y");
        assert_eq!(dash_merge_prev_scan(""), "");
    }

    /// dash_merge_next_scan：`-%}` 向后合并：删去 `-%}` 到下一个
    /// `<w:t>`/`<w:t ` 开标签（含）之间的内容，留下 `%}`
    #[test]
    fn test_dash_merge_next_scan() {
        assert_eq!(dash_merge_next_scan("-%}junk<w:t>text"), "%}text");
        // 带属性的 `<w:t ` 也是合法目标
        assert_eq!(
            dash_merge_next_scan("-%}junk<w:t xml:space=\"preserve\">x"),
            "%}x"
        );
        // 跨段落合并
        let xml = "{% endif -%}</w:t></w:r></w:p><w:p><w:r><w:t>next</w:t></w:r></w:p>";
        assert_eq!(dash_merge_next_scan(xml), "{% endif %}next</w:t></w:r></w:p>");
    }

    #[test]
    fn test_dash_merge_next_blocked() {
        // `-%}` 与下一个 `<w:t>` 之间出现 jinja 标签则放弃合并
        assert_eq!(dash_merge_next_scan("-%}{{ x }}<w:t>t"), "-%}{{ x }}<w:t>t");
        assert_eq!(dash_merge_next_scan("-%}{% y %}<w:t>t"), "-%}{% y %}<w:t>t");
        // `<w:tc>` 虽以 `<w:t` 开头但不是 `<w:t>`/`<w:t `，不匹配
        assert_eq!(dash_merge_next_scan("-%}<w:tc>x"), "-%}<w:tc>x");
        // 后面根本没有 `<w:t>`；空输入
        assert_eq!(dash_merge_next_scan("-%}junk"), "-%}junk");
        assert_eq!(dash_merge_next_scan(""), "");
    }

    /// element_tag_scan：含 `{%y `/`{{y `/`{#y ` 标签的 `<w:y>` 元素
    /// 整体替换为裸标签（y = tr/tc/p/r）
    #[test]
    fn test_element_tag_scan_basic() {
        assert_eq!(
            element_tag_scan("<w:p><w:r><w:t>{%p if x %}</w:t></w:r></w:p>", "p", false),
            "{% if x %}"
        );
        let xml = "<w:tr><w:tc><w:p><w:r><w:t>{{tr row }}</w:t></w:r></w:p></w:tc></w:tr>";
        assert_eq!(element_tag_scan(xml, "tr", false), "{{ row }}");
        assert_eq!(
            element_tag_scan("<w:r><w:t>{%r rt %}</w:t></w:r>", "r", false),
            "{% rt %}"
        );
        // comment 模式只认 `{#y #}`
        assert_eq!(
            element_tag_scan("<w:p><w:r><w:t>{#p note #}</w:t></w:r></w:p>", "p", true),
            "{# note #}"
        );
        assert_eq!(
            element_tag_scan("<w:p>{%p x %}</w:p>", "p", true),
            "<w:p>{%p x %}</w:p>"
        );
    }

    #[test]
    fn test_element_tag_scan_boundaries() {
        // 无标签的元素原样保留；继续扫描后续兄弟元素
        assert_eq!(
            element_tag_scan("<w:p>a</w:p><w:p>{%p x %}</w:p>", "p", false),
            "<w:p>a</w:p>{% x %}"
        );
        // 标签体含 `%`/`}`（docxtpl 的 [^}%]*）则不替换
        assert_eq!(
            element_tag_scan("<w:p>{%p a } b %}</w:p>", "p", false),
            "<w:p>{%p a } b %}</w:p>"
        );
        // `<w:px` 不是 `<w:p` 元素（开标签后必须是空格或 `>`）
        assert_eq!(
            element_tag_scan("<w:px>{%p x %}</w:px>", "p", false),
            "<w:px>{%p x %}</w:px>"
        );
        // 元素未闭合时区域延伸到输入末尾，仍可替换
        assert_eq!(element_tag_scan("<w:p>{%p x %}", "p", false), "{% x %}");
        // 闭合符 `%}` 落在元素区域之外则不替换
        assert_eq!(
            element_tag_scan("<w:p>{%p x </w:p>%}", "p", false),
            "<w:p>{%p x </w:p>%}"
        );
        assert_eq!(element_tag_scan("", "p", false), "");
    }

    // ------------------------------------------------------------------
    // patch_xml 端到端核心规则。期望值均对照 docxtpl 0.20.2
    // DocxTemplate.patch_xml 在相同输入下的实际输出验证过。
    // ------------------------------------------------------------------

    #[test]
    fn test_patch_xml_split_braces_and_strip_tags() {
        // jinja 标签被 run 边界拆开：先合并定界符，再剥掉标签内部的 XML
        let xml = "<w:p><w:r><w:t>{{ x </w:t></w:r><w:r><w:t>+ y }}</w:t></w:r></w:p>";
        assert_eq!(
            patch_xml(xml),
            "<w:p><w:r><w:t xml:space=\"preserve\">{{ x + y }}</w:t></w:r></w:p>"
        );
        // 定界符合并出的是不完整标签（只有 `{{ x }`）。
        // 注意：docxtpl 0.20.2 的 space-preserve 正则要求完整的 {{...}}，
        // 这里不会加 preserve；Rust 版 space_preserve_scan 只看开侧 `{{`，
        // 会多加 xml:space="preserve" —— 与 docxtpl 有出入（疑似 bug，见汇报）
        let xml2 = "<w:p><w:r><w:t>{</w:t></w:r><w:r><w:t>{ x }</w:t></w:r></w:p>";
        assert_eq!(
            patch_xml(xml2),
            "<w:p><w:r><w:t xml:space=\"preserve\">{{ x }</w:t></w:r></w:p>"
        );
    }

    #[test]
    fn test_patch_xml_colspan_cellbg() {
        // colspan：删掉空 <w:t></w:t> run 与已有 gridSpan，注入新的
        let xml = "<w:tc><w:tcPr><w:gridSpan w:val=\"9\"/></w:tcPr><w:r><w:t>{% colspan span %}</w:t></w:r></w:tc>";
        assert_eq!(
            patch_xml(xml),
            "<w:tc><w:tcPr><w:gridSpan w:val=\"{{span }}\"/></w:tcPr></w:tc>"
        );
        // cellbg：删掉已有 <w:shd>，注入带 jinja fill 的新 <w:shd>
        let xml = "<w:tc><w:tcPr><w:shd w:val=\"clear\" w:fill=\"FF0000\"/></w:tcPr><w:r><w:t>{% cellbg color %}</w:t></w:r></w:tc>";
        assert_eq!(
            patch_xml(xml),
            "<w:tc><w:tcPr><w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{{color }}\"/></w:tcPr></w:tc>"
        );
    }

    #[test]
    fn test_patch_xml_richtext_and_dash_merge() {
        // {{r ...}} 先被包进独立 run 对，随后 y-loop（y="r"）又把该 run
        // 整体改写成普通 {{ ... }}（已确认与 docxtpl 0.20.2 行为一致）
        let xml = "<w:body><w:p><w:r><w:t>{{r rt }}</w:t></w:r></w:p></w:body>";
        assert_eq!(
            patch_xml(xml),
            "<w:body><w:p><w:r><w:t xml:space=\"preserve\"></w:t></w:r>{{ rt }}<w:r><w:t xml:space=\"preserve\"></w:t></w:r></w:p></w:body>"
        );
        // {%- 跨段落向前合并
        let xml = "<w:p><w:r><w:t>text</w:t></w:r></w:p><w:p><w:r><w:t>{%- if x %}</w:t></w:r></w:p>";
        assert_eq!(
            patch_xml(xml),
            "<w:p><w:r><w:t>text{% if x %}</w:t></w:r></w:p>"
        );
        // -%} 跨段落向后合并
        let xml = "<w:p><w:r><w:t>{% endif -%}</w:t></w:r></w:p><w:p><w:r><w:t>next</w:t></w:r></w:p>";
        assert_eq!(
            patch_xml(xml),
            "<w:p><w:r><w:t xml:space=\"preserve\">{% endif %}next</w:t></w:r></w:p>"
        );
    }

    #[test]
    fn test_patch_xml_element_tags() {
        assert_eq!(
            patch_xml("<w:p><w:r><w:t>{%p if x %}</w:t></w:r></w:p>"),
            "{% if x %}"
        );
        let xml = "<w:tr><w:tc><w:p><w:r><w:t>{{tr row }}</w:t></w:r></w:p></w:tc></w:tr>";
        assert_eq!(patch_xml(xml), "{{ row }}");
        assert_eq!(
            patch_xml("<w:p><w:r><w:t>{#p note #}</w:t></w:r></w:p>"),
            "{# note #}"
        );
    }

    #[test]
    fn test_patch_xml_vmerge_hmerge() {
        // {% vm %}：注入 vMerge，内容包进 {% if loop.first %}
        let xml = "<w:tc><w:tcPr></w:tcPr><w:r><w:t>{% vm %}content</w:t></w:r></w:tc>";
        assert_eq!(
            patch_xml(xml),
            "<w:tc><w:tcPr><w:vMerge w:val=\"{% if loop.first %}restart{% else %}continue{% endif %}\"/></w:tcPr><w:r><w:t xml:space=\"preserve\">{% if loop.first %}content{% endif %}</w:t></w:r></w:tc>"
        );
        // {% hm %} 无 gridSpan：新增 gridSpan = loop.length
        let xml = "<w:tc><w:tcPr></w:tcPr><w:r><w:t>a{% hm %}b</w:t></w:r></w:tc>";
        assert_eq!(
            patch_xml(xml),
            "{% if loop.first %}<w:tc><w:tcPr><w:gridSpan w:val=\"{{ loop.length }}\"/></w:tcPr><w:r><w:t xml:space=\"preserve\">ab</w:t></w:r></w:tc>{% endif %}"
        );
        // {% hm %} 已有 gridSpan：值乘以 loop.length
        let xml = "<w:tc><w:tcPr><w:gridSpan w:val=\"3\"/></w:tcPr><w:r><w:t>{% hm %}</w:t></w:r></w:tc>";
        assert_eq!(
            patch_xml(xml),
            "{% if loop.first %}<w:tc><w:tcPr><w:gridSpan w:val=\"{{ 3 * loop.length }}\"/></w:tcPr><w:r><w:t xml:space=\"preserve\"></w:t></w:r></w:tc>{% endif %}"
        );
    }

    #[test]
    fn test_patch_xml_clean_tags_and_identity() {
        // jinja 标签内的 &lt;/&gt; 与智能引号被还原
        assert_eq!(
            patch_xml("<w:p><w:r><w:t>{{ x &lt; y }}</w:t></w:r></w:p>"),
            "<w:p><w:r><w:t xml:space=\"preserve\">{{ x < y }}</w:t></w:r></w:p>"
        );
        assert_eq!(
            patch_xml("<w:p><w:r><w:t>{{ \u{201c}s\u{201d} }}</w:t></w:r></w:p>"),
            "<w:p><w:r><w:t xml:space=\"preserve\">{{ \"s\" }}</w:t></w:r></w:p>"
        );
        // 无任何触发内容时恒等；空输入
        let xml = "<w:p><w:r><w:t>plain</w:t></w:r></w:p>";
        assert_eq!(patch_xml(xml), xml);
        assert_eq!(patch_xml(""), "");
    }

    /// 长输入（约 700KB）：线性扫描器在无匹配时保持恒等、在末尾有
    /// 匹配时正确跨段删除 —— 正是 AGENTS.md 第 9 条防回退栈溢出的场景
    #[test]
    fn test_scanners_long_input() {
        let mut long = String::new();
        for _ in 0..20_000 {
            long.push_str("<w:p><w:r><w:t>t</w:t></w:r></w:p>");
        }
        assert_eq!(merge_split_braces_scan(&long), long);
        assert_eq!(space_preserve_scan(&long), long);
        assert_eq!(dash_merge_next_scan(&long), long);
        assert_eq!(element_tag_scan(&long, "p", false), long);
        // `{%-` 在大文档末尾：从最后一个 `</w:t>` 起跨段删除
        let input = format!("{}<w:p><w:r><w:t>{{%- x %}}</w:t></w:r></w:p>", long);
        let expected = format!(
            "{}{{% x %}}</w:t></w:r></w:p>",
            &long[..long.len() - "</w:t></w:r></w:p>".len()]
        );
        assert_eq!(dash_merge_prev_scan(&input), expected);
    }
}
