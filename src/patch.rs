//! Port of docxtpl's patch_xml / resolve_listing logic (regex-based XML cleaning).

use fancy_regex::{Captures, Regex};
use std::borrow::Cow;
use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

/// Timing instrumentation, enabled with PATCH_TIMING=1.
fn patch_timing_enabled() -> bool {
    // env::var scans the environment and allocates; cache the answer
    static ON: std::sync::OnceLock<bool> = std::sync::OnceLock::new();
    *ON.get_or_init(|| std::env::var("PATCH_TIMING").is_ok())
}

fn patch_debug_enabled() -> bool {
    static ON: std::sync::OnceLock<bool> = std::sync::OnceLock::new();
    *ON.get_or_init(|| std::env::var("PATCH_DEBUG").is_ok())
}

macro_rules! timed {
    ($name:expr, $e:expr) => {{
        let __t0 = std::time::Instant::now();
        let __r = $e;
        if patch_timing_enabled() {
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
    if patch_debug_enabled() {
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
/// Single-sweep replacement for `contains_any(xml, &["{<","%<","}<","#<"])`:
/// a match is a '<' directly preceded by a jinja delimiter byte. 4 memmem
/// sweeps collapse into one SIMD memchr scan.
fn gate_delim_tag(xml: &str) -> bool {
    let b = xml.as_bytes();
    let mut i = 0usize;
    while let Some(rel) = memchr::memchr(b'<', &b[i..]) {
        let p = i + rel;
        if p > 0 && matches!(b[p - 1], b'{' | b'%' | b'}' | b'#') {
            return true;
        }
        i = p + 1;
    }
    false
}

/// Single-sweep replacement for `contains_any(xml, &["{{","{%","{#"])`:
/// a match is a '{' directly followed by one of '{','%','#'.
fn gate_jinja_open(xml: &str) -> bool {
    let b = xml.as_bytes();
    let mut i = 0usize;
    while let Some(rel) = memchr::memchr(b'{', &b[i..]) {
        let p = i + rel;
        if matches!(b.get(p + 1), Some(b'{') | Some(b'%') | Some(b'#')) {
            return true;
        }
        i = p + 1;
    }
    false
}

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
        // jump to the next candidate delimiter byte (SIMD memchr): open side
        // '{' ; close side one of '%','}','#'
        let next = match (
            memchr::memchr3(b'{', b'%', b'#', &b[i..]),
            memchr::memchr(b'}', &b[i..]),
        ) {
            (Some(a), Some(bb)) => i + a.min(bb),
            (Some(a), None) => i + a,
            (None, Some(bb)) => i + bb,
            (None, None) => break,
        };
        i = next;
        let c = b[i];
        // open side: '{' ... one of '{','%','#' ; close side: one of '%','}','#' ... '}'
        let open_side = c == b'{';
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
        // no group references: skip the 18 full-string replace scans
        if !replacement.contains('$') {
            return replacement.clone();
        }
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
pub fn decode_text_entities(xml: &str) -> Cow<'_, str> {
    if !xml.contains('&') {
        return Cow::Borrowed(xml);
    }
    // hand-rolled equivalent of the `(?<=>)([^<>]*)(?=<)` pass: decode
    // entities in text nodes only; borrowed when nothing actually changed
    // (e.g. only `&lt;`/`&amp;` which are kept escaped)
    let b = xml.as_bytes();
    let mut out: Option<String> = None;
    let mut copied = 0usize;
    let mut i = 0usize;
    while i < b.len() {
        if b[i] == b'>' {
            let text_start = i + 1;
            let mut j = text_start;
            while j < b.len() && b[j] != b'<' && b[j] != b'>' {
                j += 1;
            }
            if j < b.len() && b[j] == b'<' && j > text_start {
                let text = &xml[text_start..j];
                if text.contains('&') {
                    let decoded = decode_entities_keep_markup(text);
                    if decoded != text {
                        let o = out.get_or_insert_with(|| String::with_capacity(xml.len()));
                        o.push_str(&xml[copied..text_start]);
                        o.push_str(&decoded);
                        copied = j;
                    }
                }
            }
            i = j; // j >= text_start > i, always forward
            continue;
        }
        i += 1;
    }
    match out {
        Some(mut o) => {
            o.push_str(&xml[copied..]);
            Cow::Owned(o)
        }
        None => Cow::Borrowed(xml),
    }
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
pub fn patch_xml(src_xml: &str) -> Cow<'_, str> {
    // replace {<something>{ by {{   ( works with {{ }} {% and %} {# and #})
    // (hand-rolled linear scan; fancy_regex spends ~20ms/500KB here even when
    // nothing matches)
    // gate: a match requires a delimiter directly followed by a tag
    let mut xml: Cow<'_, str> = if gate_delim_tag(src_xml) {
        Cow::Owned(timed!("merge_split_braces", merge_split_braces_scan(src_xml)))
    } else {
        Cow::Borrowed(src_xml)
    };
    // apply a pass, going (or staying) owned
    macro_rules! pass {
        ($name:expr, $e:expr) => {
            xml = Cow::Owned(timed!($name, $e))
        };
    }

    // replace {{<some tags>jinja2 stuff<some other tags>}} by {{jinja2 stuff}},
    // ensure space preservation for runs containing jinja tags, and unescape
    // entities/smart quotes inside jinja tags — fused into a single linear
    // scan (fused_jinja_scan) that produces byte-identical output to running
    // strip_tags_in_jinja_scan + space_preserve_scan + clean_tags_scan in
    // sequence, with one full-document allocation instead of three.
    // gate: a match in any of the three passes requires a jinja open marker
    // (the space-preserve pass additionally requires `<w:t>`, which the fused
    // scan handles internally: without one, nothing is rewritten)
    //
    // Exception: docxtpl runs colspan/cellbg BETWEEN strip_tags_in_jinja and
    // space_preserve (their empty-run removal matches `<w:t></w:t>`, which the
    // space pass would already have rewritten). A fused pass cannot reproduce
    // that interleaving, so templates with those (rare) markers keep the
    // original three-pass sequence below.
    let mut cell_directive = xml.contains("colspan") || xml.contains("cellbg");
    // directive gates precomputed on the post-jinja-pass string (saves the
    // re-scan in the colspan/cellbg branches below)
    let mut gate_colspan: Option<bool> = None;
    let mut gate_cellbg: Option<bool> = None;
    if gate_jinja_open(&xml) {
        if cell_directive {
            pass!("strip_tags_in_jinja", strip_tags_in_jinja_scan(&xml));
        } else {
            let fused = timed!("fused_jinja", fused_jinja_scan(&xml));
            // compute both directive gates in one shot: the gate pair below
            // re-checks the same string
            let fused_colspan = fused.contains("colspan");
            let fused_cellbg = fused.contains("cellbg");
            if fused_colspan || fused_cellbg {
                // pathological: the marker was formed across a stripped run
                // split — redo with the original pass order so colspan/cellbg
                // still run between strip_tags_in_jinja and space_preserve
                pass!("strip_tags_in_jinja", strip_tags_in_jinja_scan(&xml));
                cell_directive = true;
            } else {
                xml = Cow::Owned(fused);
                gate_colspan = Some(false);
                gate_cellbg = Some(false);
            }
        }
    }

    // manage table cell colspan
    if gate_colspan.unwrap_or_else(|| xml.contains("colspan")) {
    pass!("colspan", sub(
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
    // (valid when the colspan gate above was precomputed false: the colspan
    // pass then skipped and the string is unchanged)
    if gate_cellbg.unwrap_or_else(|| xml.contains("cellbg")) {
    pass!("cellbg", sub(
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

    // ensure space preservation (see the fused_jinja note above)
    if cell_directive && xml.contains("<w:t>") && contains_any(&xml, &["{{", "{%"]) {
        pass!("space_preserve", space_preserve_scan(&xml));
    }
    if contains_any(&xml, &["{{r", "{%r"]) {
        pass!("richtext_tag", sub(
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
        pass!("dash_merge_prev", dash_merge_prev_scan(&xml));
    }
    // -%} will merge with next paragraph text
    if xml.contains("-%}") {
        pass!("dash_merge_next", dash_merge_next_scan(&xml));
    }

    // replace into xml code the row/paragraph/run containing
    // {%y xxx %} / {{y xxx}} / {#y xxx #} template tag by the tag alone
    // without any surrounding <w:y> tags (hand-rolled linear scan; the
    // original tempered-dot regex overflows the backtrack stack on large
    // documents). Marker positions are collected with aho-corasick sweeps;
    // a pass that replaces elements SHIFTS all later positions, so the
    // remaining buckets are re-collected on the new string (matches the
    // original per-y gate + collect-on-current-string semantics).
    {
        const YS: [&str; 4] = ["tr", "tc", "p", "r"];
        let pats: Vec<String> = YS
            .iter()
            .flat_map(|y| [format!("{{{{{} ", y), format!("{{%{} ", y)])
            .chain(YS[..3].iter().map(|y| format!("{{#{} ", y)))
            .collect();
        let ac = aho_corasick::AhoCorasick::new(&pats).unwrap();
        let mut buckets: [Vec<usize>; 11] = Default::default();
        fn collect(xml: &str, ac: &aho_corasick::AhoCorasick, from: usize, buckets: &mut [Vec<usize>; 11]) {
            for b in buckets.iter_mut().skip(from) {
                b.clear();
            }
            for m in ac.find_iter(xml.as_bytes()) {
                let pi = m.pattern().as_usize();
                if pi >= from {
                    buckets[pi].push(m.start());
                }
            }
        }
        collect(&xml, &ac, 0, &mut buckets);
        for (i, y) in YS.iter().enumerate() {
            let mut pos = std::mem::take(&mut buckets[2 * i]);
            pos.extend_from_slice(&buckets[2 * i + 1]);
            if !pos.is_empty() {
                pos.sort_unstable();
                let (new, changed) = element_tag_scan_hits(&xml, y, false, &pos);
                xml = Cow::Owned(new);
                if changed {
                    collect(&xml, &ac, 2 * i + 2, &mut buckets);
                }
            }
        }
        for (j, y) in YS[..3].iter().enumerate() {
            if !buckets[8 + j].is_empty() {
                let (new, changed) = element_tag_scan_hits(&xml, y, true, &buckets[8 + j]);
                xml = Cow::Owned(new);
                if changed {
                    collect(&xml, &ac, 8 + j + 1, &mut buckets);
                }
            }
        }
    }

    // add vMerge
    // use {% vm %} to make this table cell and its copies
    // be vertically merged within a {% for %}
    if xml.contains("{%") && xml.contains("vm") {
    pass!("vmerge", sub(
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
    pass!("hmerge", sub(
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
    // (see the fused_jinja note above)
    if cell_directive && contains_any(&xml, &["{{", "{%"]) {
        pass!("clean_tags", clean_tags_scan(&xml));
    }

    xml
}

/// replace only first occurrence
fn sub_first(pattern: &str, replacement: &str, text: &str) -> String {
    let rex = re(pattern);
    rex.replace(text, replacement).to_string()
}

/// Port of DocxTemplate.resolve_listing
pub fn resolve_listing(xml: &str) -> Cow<'_, str> {
    // resolve_text only rewrites \t, \n, \x07 and \x0c; without any of them
    // the whole pass is the identity, so skip the full-document copy
    if !xml.contains(['\t', '\n', '\u{7}', '\u{c}']) {
        return Cow::Borrowed(xml);
    }
    Cow::Owned(timed!("resolve_listing_scan", resolve_listing_scan(xml)))
}

/// Find the next `<tag>`/`<tag ...>` opening tag (regex `<tag(?: [^>]*)?>`,
/// which notably does NOT match longer tags like `<w:pPr>`/`<w:rPr>`/
/// `<w:tab/>`): returns (start, end_after_gt).
fn find_open_tag(s: &str, tag: &str) -> Option<(usize, usize)> {
    let mut i = 0usize;
    while let Some(rel) = s[i..].find(tag) {
        let p = i + rel;
        let after = p + tag.len();
        match s.as_bytes().get(after) {
            Some(b'>') => return Some((p, after + 1)),
            Some(b' ') => {
                // [^>]* up to the first '>'; without one no match is
                // possible here or anywhere further
                let gt = s[after..].find('>')?;
                return Some((p, after + gt + 1));
            }
            _ => i = p + 1,
        }
    }
    None
}

/// first `<open>...</close>` (non-greedy), like regex find on `<open>.*?</close>`
fn find_first_elem<'a>(s: &'a str, open: &str, close: &str) -> Option<&'a str> {
    let start = s.find(open)?;
    let end = s[start + open.len()..]
        .find(close)
        .map(|i| start + open.len() + i + close.len())?;
    Some(&s[start..end])
}

/// Hand-rolled linear equivalent of the fancy-regex paragraph/run/w:t
/// scans (the backtracking VM dominated render time for listing-heavy
/// templates): for each `<w:p>` containing a listing char, rewrite \t, \n,
/// \x07, \x0c inside its `<w:t>` elements.
///
/// Marker-first: listing chars are rare, so we locate them with SIMD memchr
/// and resolve only the enclosing paragraph (nearest valid `<w:p` open whose
/// region, up to the first `</w:p>`, contains the char — `<w:p>` never
/// nests). Paragraphs without listing chars are bulk-copied instead of being
/// walked one by one. Chars outside any paragraph (e.g. newlines left
/// between paragraphs by the jinja newline pass) pass through verbatim,
/// exactly like the sequential scan's gap copying.
fn resolve_listing_scan(xml: &str) -> String {
    let b = xml.as_bytes();
    let n = b.len();
    let mut out = String::with_capacity(n);
    let mut copied = 0usize; // xml[..copied] already flushed to out
    let mut i = 0usize;
    while i < n {
        // next listing char (SIMD memchr): \t \n \x07 \x0c
        let next = match (
            memchr::memchr3(b'\t', b'\n', 0x07, &b[i..]),
            memchr::memchr(0x0c, &b[i..]),
        ) {
            (Some(a), Some(bb)) => i + a.min(bb),
            (Some(a), None) => i + a,
            (None, Some(bb)) => i + bb,
            (None, None) => break,
        };
        // enclosing paragraph: the sequential scan processes opens left to
        // right, so the char's paragraph is the LEFTMOST valid open (at or
        // after `copied`) whose region — up to the first `</w:p>` — contains
        // it. For opens o1 < o2 the first close after o1 is <= the first
        // close after o2, so the enclosing opens form a chain; walk it
        // backwards from the nearest open to find the leftmost member.
        let mut end = next;
        let mut para: Option<(usize, usize)> = None; // (open_start, region_end)
        loop {
            let mut open = None;
            while let Some(ps) = xml[..end].rfind("<w:p") {
                let after = ps + 4;
                if matches!(b.get(after), Some(b' ') | Some(b'>')) {
                    open = Some((ps, after));
                    break;
                }
                end = ps; // e.g. `<w:pPr`: not an opening of w:p
            }
            let Some((ps, after)) = open else { break };
            if ps < copied {
                break; // already consumed by the sequential walk
            }
            let Some(close) = xml[after..].find("</w:p>") else {
                // without a close no later paragraph can have one either:
                // everything left passes through verbatim
                out.push_str(&xml[copied..]);
                return out;
            };
            let para_end = after + close + "</w:p>".len();
            if para_end <= next {
                break; // nearest open does not enclose the char: no open does
            }
            para = Some((ps, para_end));
            end = ps;
        }
        match para {
            Some((ps, para_end)) => {
                // rewrite the paragraph, bulk-copy everything before it
                out.push_str(&xml[copied..ps]);
                resolve_paragraph_scan(&xml[ps..para_end], &mut out);
                copied = para_end;
                i = para_end;
            }
            None => {
                i = next + 1; // char outside any paragraph: verbatim
            }
        }
    }
    out.push_str(&xml[copied..]);
    out
}

fn resolve_paragraph_scan(whole: &str, out: &mut String) {
    let ppr = find_first_elem(whole, "<w:pPr>", "</w:pPr>").unwrap_or("");
    let mut rest = whole;
    while let Some((rs, re)) = find_open_tag(rest, "<w:r") {
        let Some(close) = rest[re..].find("</w:r>") else { break };
        let run_end = re + close + "</w:r>".len();
        out.push_str(&rest[..rs]);
        resolve_run_scan(&rest[rs..run_end], ppr, out);
        rest = &rest[run_end..];
    }
    out.push_str(rest);
}

fn resolve_run_scan(run: &str, ppr: &str, out: &mut String) {
    let rpr = find_first_elem(run, "<w:rPr>", "</w:rPr>").unwrap_or("");
    let mut rest = run;
    while let Some((ts, te)) = find_open_tag(rest, "<w:t") {
        let Some(close) = rest[te..].find("</w:t>") else { break };
        let elem_end = te + close + "</w:t>".len();
        out.push_str(&rest[..ts]);
        resolve_text_into(&rest[ts..elem_end], rpr, ppr, out);
        rest = &rest[elem_end..];
    }
    out.push_str(rest);
}

/// resolve_text on one `<w:t>...</w:t>` element. Keeps the original
/// sequential str::replace chain (the replacements are rescanned by the
/// later replaces), but only on the small element string.
fn resolve_text_into(s: &str, run_properties: &str, paragraph_properties: &str, out: &mut String) {
    if !s.contains(['\t', '\n', '\u{7}', '\u{c}']) {
        out.push_str(s);
        return;
    }
    let s = s.replace(
        '\t',
        &format!(
            "</w:t></w:r><w:r>{}<w:tab/></w:r><w:r>{}<w:t xml:space=\"preserve\">",
            run_properties, run_properties
        ),
    );
    let s = s.replace(
        '\u{7}',
        &format!(
            "</w:t></w:r></w:p><w:p>{}<w:r>{}<w:t xml:space=\"preserve\">",
            paragraph_properties, run_properties
        ),
    );
    let s = s.replace('\n', "</w:t><w:br/><w:t xml:space=\"preserve\">");
    let s = s.replace(
        '\u{c}',
        &format!(
            "</w:t></w:r></w:p><w:p><w:r><w:br w:type=\"page\"/></w:r></w:p><w:p>{}<w:r>{}<w:t xml:space=\"preserve\">",
            paragraph_properties, run_properties
        ),
    );
    out.push_str(&s);
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
///
/// Marker-first implementation: jinja markers (`{%tr ` etc.) are extremely
/// rare compared to `<w:y` elements, so we locate markers first (memmem) and
/// resolve only the elements that enclose them, instead of probing every
/// element opening in the document.
///
/// Exact equivalence with the element-first scan (see git history): the
/// original walk processes element openings left to right; an element whose
/// region (up to its first close tag) contains a marker is checked against
/// the EARLIEST marker in that region; on success the whole region is
/// replaced by the bare tag and scanning resumes after it; on failure
/// scanning resumes just after the opening tag, so the same marker is tried
/// again against the next enclosing open. For a marker at position p the
/// opens that enclose it form a chain o1 < o2 < ... < ok (for o < o' the
/// first close after o is <= the first close after o'), and the element-first
/// walk tries them leftmost-first. The code below collects that chain by
/// walking backwards from p, then replays the same checks in the same order.
pub fn element_tag_scan(xml: &str, y: &str, comment: bool) -> String {
    // collect marker occurrences (rare) in document order
    let m_var = format!("{{{{{} ", y);
    let m_stmt = format!("{{%{} ", y);
    let m_comment = format!("{{#{} ", y);
    let pats: &[String] = if comment {
        std::slice::from_ref(&m_comment)
    } else {
        &[m_var, m_stmt]
    };
    let mut positions: Vec<usize> = Vec::new();
    for pat in pats {
        let mut from = 0usize;
        while let Some(rel) = xml[from..].find(pat.as_str()) {
            positions.push(from + rel);
            from += rel + 1;
        }
    }
    if positions.is_empty() {
        return xml.to_string();
    }
    positions.sort_unstable();
    element_tag_scan_hits(xml, y, comment, &positions).0
}

/// element_tag_scan with pre-collected (sorted) marker positions: patch_xml
/// gathers every y's markers in a single aho-corasick sweep and calls this
/// directly, skipping both the per-y gate sweeps and the collection scans.
/// Returns the new string and whether any element was actually replaced
/// (callers that chain multiple y passes must re-collect positions then).
fn element_tag_scan_hits(xml: &str, y: &str, comment: bool, positions: &[usize]) -> (String, bool) {
    let open_prefix = format!("<w:{}", y);
    let close_tag = format!("</w:{}>", y);
    // `{%y `/`{{y `/`{#y `: two delimiter bytes + y + space
    let marker_len = y.len() + 3;

    let mut out = String::with_capacity(xml.len());
    let mut copied = 0usize; // everything before `copied` already emitted
    let mut i = 0usize; // element-first scan position lower bound
    let mut any_replaced = false;

    for &p in positions {
        if p < i {
            continue; // inside an already-consumed region or skipped open tag
        }
        // marker kind from its second byte ('{' / '%' / '#')
        let (close_tok, forbidden): (&str, &[char]) = if comment {
            ("#}", &['#', '}'])
        } else if xml.as_bytes().get(p + 1) == Some(&b'{') {
            ("}}", &['%', '}'])
        } else {
            ("%}", &['%', '}'])
        };
        // chain of valid opens enclosing p: nearest-first walk backwards
        let mut chain: Vec<(usize, usize, usize)> = Vec::new(); // (open, after_open, region_end)
        let mut cursor = p;
        loop {
            // nearest valid `<w:y` open before `cursor`
            let mut end = cursor;
            let open = loop {
                let Some(o) = xml[..end].rfind(&open_prefix) else {
                    break None;
                };
                let after = o + open_prefix.len();
                if matches!(xml.as_bytes().get(after), Some(b' ') | Some(b'>')) {
                    break Some(o);
                }
                end = o; // e.g. `<w:pPr`: not an opening of w:p
            };
            let Some(open) = open else { break };
            if open < i {
                break; // already consumed/skipped by the element-first walk
            }
            let after_open = open + open_prefix.len();
            let region_end = xml[after_open..]
                .find(&close_tag)
                .map(|c| after_open + c + close_tag.len())
                .unwrap_or(xml.len());
            if region_end <= p {
                break; // nearest open does not enclose p: no open does
            }
            chain.push((open, after_open, region_end));
            cursor = open;
        }
        // try the chain leftmost-first, like the element-first walk
        let mut handled = false;
        for &(open, _after_open, region_end) in chain.iter().rev() {            let region = &xml[open..region_end];
            let after_marker = p - open + marker_len;
            // close token must appear before any '%'/'}' (statement) or
            // '#'/'}' (comment), matching [^}%]* / [^}#]* of the regex
            let inner = region[after_marker..]
                .find(close_tok)
                .map(|close_rel| &region[after_marker..after_marker + close_rel]);
            let ok = inner
                .map(|inner| !inner.chars().any(|c| forbidden.contains(&c)))
                .unwrap_or(false);
            if ok {
                // emit: everything before the element + bare tag; skip region
                out.push_str(&xml[copied..open]);
                out.push_str(&xml[p..p + 2]);
                out.push(' ');
                out.push_str(inner.unwrap());
                out.push_str(close_tok);
                copied = region_end;
                i = region_end;
                handled = true;
                any_replaced = true;
                break;
            }
        }
        if !handled {
            // whole chain failed: the walk resumes just after the nearest
            // failed opening tag (markers inside it are skipped via p < i)
            if let Some(&(_open, after_open, _)) = chain.first() {
                i = after_open;
            }
        }
    }
    out.push_str(&xml[copied..]);
    (out, any_replaced)
}

/// Hand-rolled linear equivalent of the strip_tags_in_jinja pass:
/// `(?s)\{%(?:(?!%\}).)*|\{#(?:(?!#\}).)*|\{\{(?:(?!}\}).)*` where each match
/// has its inner `</w:t>.*?(<w:t>|<w:t [^>]*>)` runs removed (jinja tags
/// split across runs by Word). The tempered dot is greedy up to the FIRST
/// closing delimiter (or EOF when there is none — the closer is NOT part of
/// the pattern, so the match succeeds either way and does not consume the
/// closer). The fancy-regex pair cost ~half of patch_xml on large documents.
fn strip_tags_in_jinja_scan(xml: &str) -> String {
    let mut out = String::with_capacity(xml.len());
    let mut rest = xml;
    loop {
        let b = rest.as_bytes();
        if b.len() < 2 {
            break;
        }
        // earliest opener among {% {# {{
        let mut o = None;
        let mut i = 0usize;
        while let Some(rel) = memchr::memchr(b'{', &b[i..]) {
            let p = i + rel;
            if p + 1 >= b.len() {
                break;
            }
            if matches!(b[p + 1], b'{' | b'%' | b'#') {
                o = Some(p);
                break;
            }
            i = p + 1;
        }
        let Some(o) = o else { break };
        let closer = match b[o + 1] {
            b'%' => "%}",
            b'#' => "#}",
            _ => "}}",
        };
        let content_start = o + 2;
        let content_end = rest[content_start..]
            .find(closer)
            .map(|r| content_start + r)
            .unwrap_or(rest.len());
        out.push_str(&rest[..content_start]);
        strip_wt_pairs(&rest[content_start..content_end], &mut out);
        rest = &rest[content_end..];
    }
    out.push_str(rest);
    out
}

/// Remove every `</w:t>.*?(<w:t>|<w:t [^>]*>)` run (non-greedy, like the
/// original inner regex, so each removal spans up to the NEAREST following
/// `<w:t>` opening) from a jinja tag's content, appending the result.
fn strip_wt_pairs(s: &str, out: &mut String) {
    if !s.contains("</w:t>") {
        out.push_str(s);
        return;
    }
    let b = s.as_bytes();
    let mut copied = 0usize;
    let mut i = 0usize;
    while let Some(rel) = s[i..].find("</w:t>") {
        let close_start = i + rel;
        // earliest "<w:t>" or "<w:t ...>" opening tag after the close tag
        let mut j = close_start + 6;
        let open_end = loop {
            if j + 5 > b.len() {
                break None;
            }
            if b[j] == b'<' && b[j + 1..].starts_with(b"w:t") {
                if b[j + 4] == b'>' {
                    break Some(j + 5);
                }
                if b[j + 4] == b' ' {
                    // `<w:t [^>]*>`: up to and including the first '>'
                    match s[j + 5..].find('>') {
                        Some(g) => break Some(j + 5 + g + 1),
                        // no '>' left anywhere: no later match can complete
                        None => break None,
                    }
                }
            }
            j += 1;
        };
        match open_end {
            Some(e) => {
                out.push_str(&s[copied..close_start]);
                copied = e;
                i = e;
            }
            None => break,
        }
    }
    out.push_str(&s[copied..]);
}

/// Hand-rolled linear equivalent of the clean_tags pass:
/// `(?<=\{[\{%])(.*?)(?=[\}%]\})` — inside `{{ ... }}` / `{% ... %}` tags
/// (the closer is the first `}}` OR `%}` regardless of the opener; with no
/// closer at all, nothing matches anywhere further), unescape `&#8216;` /
/// `&lt;` / `&gt;` and normalize smart quotes. The lookbehind + lazy dot +
/// lookahead regex cost ~the other half of patch_xml on large documents.
fn clean_tags_scan(xml: &str) -> String {
    let mut out = String::with_capacity(xml.len());
    let mut rest = xml;
    loop {
        let b = rest.as_bytes();
        if b.len() < 4 {
            break; // need at least opener (2) + closer (2)
        }
        // earliest opener among {{ {%
        let mut o = None;
        let mut i = 0usize;
        while i + 3 < b.len() {
            if b[i] == b'{' && (b[i + 1] == b'{' || b[i + 1] == b'%') {
                o = Some(i);
                break;
            }
            i += 1;
        }
        let Some(o) = o else { break };
        let content_start = o + 2;
        // first "}}" or "%}" at/after content_start
        let mut j = content_start;
        let mut cend = None;
        while j + 1 < b.len() {
            if (b[j] == b'}' || b[j] == b'%') && b[j + 1] == b'}' {
                cend = Some(j);
                break;
            }
            j += 1;
        }
        match cend {
            Some(e) => {
                out.push_str(&rest[..content_start]);
                clean_tag_content(&rest[content_start..e], &mut out);
                // the closer is not consumed (lookahead), continue there
                rest = &rest[e..];
            }
            None => break, // no closer anywhere: no further matches possible
        }
    }
    out.push_str(rest);
    out
}

/// Apply the clean_tags replacements to a tag's content, appending to `out`:
/// `&#8216;` -> ', `&lt;` -> <, `&gt;` -> >, U+2018/2019 -> ', U+201C/201D
/// -> ". Single pass: the replacements introduce neither `&` nor smart
/// quotes, so the sequential `str::replace` chain is equivalent.
fn clean_tag_content(s: &str, out: &mut String) {
    if !s.contains('&') && !s.contains(['\u{2018}', '\u{2019}', '\u{201c}', '\u{201d}']) {
        out.push_str(s);
        return;
    }
    let b = s.as_bytes();
    let mut copied = 0usize;
    let mut i = 0usize;
    while i < b.len() {
        let rep: Option<(&str, usize)> = match b[i] {
            b'&' => {
                if s[i..].starts_with("&#8216;") {
                    Some(("'", "&#8216;".len()))
                } else if s[i..].starts_with("&lt;") {
                    Some(("<", 4))
                } else if s[i..].starts_with("&gt;") {
                    Some((">", 4))
                } else {
                    None
                }
            }
            // U+2018/2019/201C/201D are E2 80 98/99/9C/9D in UTF-8
            0xE2 if i + 2 < b.len() && b[i + 1] == 0x80 => match b[i + 2] {
                0x98 | 0x99 => Some(("'", 3)),
                0x9C | 0x9D => Some(("\"", 3)),
                _ => None,
            },
            _ => None,
        };
        match rep {
            Some((r, len)) => {
                out.push_str(&s[copied..i]);
                out.push_str(r);
                i += len;
                copied = i;
            }
            None => i += 1,
        }
    }
    out.push_str(&s[copied..]);
}

// --------------------------------------------------------------------
// Fused strip_tags_in_jinja + space_preserve + clean_tags (single pass)
// --------------------------------------------------------------------

const WT_PLAIN: &str = "<w:t>";
const WT_PRESERVE: &str = "<w:t xml:space=\"preserve\">";

/// Token kinds detected in the post-strip stream by the fused emitter.
#[derive(Clone, Copy)]
enum Tok {
    /// exact `<w:t>` open tag
    Wt,
    /// `{{` or `{%`
    Open,
    /// `}}` or `%}` (only relevant inside a clean region)
    Close,
}

enum Dec {
    Token(Tok, usize),
    /// the first carried byte is definitely plain text (the remainder may
    /// still start a token, e.g. `<w:t}` followed by `}`)
    Plain,
    Undecided,
}

/// Decide whether the carried bytes (a possible token prefix, <= 5 ASCII
/// bytes) already form a token. All token bytes are ASCII; a non-ASCII byte
/// can never be part of a token and is never put into the carry.
fn decide(carry: &[u8], in_clean: bool) -> Dec {
    match carry[0] {
        b'<' => {
            const WT: &[u8] = b"<w:t>";
            if carry.len() < WT.len() {
                if WT.starts_with(carry) {
                    Dec::Undecided
                } else {
                    Dec::Plain
                }
            } else if &carry[..WT.len()] == WT {
                Dec::Token(Tok::Wt, WT.len())
            } else {
                Dec::Plain
            }
        }
        b'{' => {
            if carry.len() < 2 {
                Dec::Undecided
            } else if carry[1] == b'{' || carry[1] == b'%' {
                Dec::Token(Tok::Open, 2)
            } else {
                Dec::Plain
            }
        }
        b'}' | b'%' if in_clean => {
            if carry.len() < 2 {
                Dec::Undecided
            } else if carry[1] == b'}' {
                Dec::Token(Tok::Close, 2)
            } else {
                Dec::Plain
            }
        }
        _ => Dec::Plain,
    }
}

/// Stream emitter that applies the space_preserve and clean_tags passes on
/// the post-strip text as it is produced by the strip driver, so the whole
/// strip -> space -> clean composition happens in a single scan with a
/// single full-document allocation.
///
/// Equivalence with the sequential composition, by pass:
/// - space_preserve_scan(A) rewrites an exact `<w:t>` to
///   `<w:t xml:space="preserve">` iff the text up to the next exact `<w:t>`
///   (or EOF) contains `{{`/`{%`. The emitter sees exactly the bytes of A:
///   strip-removed `</w:t>...<w:t>` pairs never enter the stream, so "next
///   `<w:t>`" is automatically the post-strip one. The decision for a seen
///   `<w:t>` stays open (`pending_wt`, following text goes to `hold`) until
///   an opener (rewrite), the next `<w:t>` (plain), or EOF (plain). When a
///   `<w:t>` sits inside a clean region whose closer arrives first, it is
///   flushed plain and its position remembered (`retro_wt`); a later opener
///   (before the next `<w:t>`) still upgrades it in place. The rewrite only
///   inserts bytes into the opening tag, so it cannot affect clean's region
///   structure (no `{`/`}`/`%`) nor its content replacements.
/// - clean_tags_scan(B) cleans regions `{{`/`{%` .. first `}}`/`%}` (opener
///   set excludes `{#`; closer not consumed; no closer anywhere -> verbatim).
///   The emitter tracks exactly those regions on the stream (`cbuf`), which
///   also covers `{{`/`{%` nested inside `{#...#}` strip regions and clean
///   closers that precede the strip closer (e.g. `{% a }} b %}`).
struct FusedEmitter {
    out: String,
    /// clean-region content accumulator: Some while inside `{{`/`{%` .. first
    /// `}}`/`%}`; cleaned with clean_tag_content when the region closes
    cbuf: Option<String>,
    /// kept buffer for the next clean region (jinja tags come in thousands;
    /// reusing one allocation avoids a fresh String per tag)
    cbuf_spare: String,
    /// undecided token prefix carried across feed() boundaries (<=4 ASCII)
    carry: Vec<u8>,
    /// a `<w:t>` was seen but not yet emitted: its space_preserve decision is
    /// still open; plain text seen meanwhile accumulates in `hold`
    pending_wt: bool,
    hold: String,
    /// position in `out` of a `<w:t>` that had to be flushed plain before its
    /// decision because its enclosing clean region closed
    retro_wt: Option<usize>,
}

impl FusedEmitter {
    fn new(cap: usize) -> Self {
        FusedEmitter {
            out: String::with_capacity(cap),
            cbuf: None,
            cbuf_spare: String::new(),
            carry: Vec::new(),
            pending_wt: false,
            hold: String::new(),
            retro_wt: None,
        }
    }

    #[inline]
    fn plain_dest(&mut self) -> &mut String {
        if self.pending_wt {
            &mut self.hold
        } else if let Some(b) = &mut self.cbuf {
            b
        } else {
            &mut self.out
        }
    }

    #[inline]
    fn push_plain(&mut self, s: &str) {
        self.plain_dest().push_str(s);
    }

    /// Emit the pending `<w:t>` (as `tag`) plus any held text. The state
    /// (inside/outside a clean region) cannot have changed while the tag was
    /// pending, because every state-changing event flushes it first.
    fn flush_pending(&mut self, tag: &str) {
        debug_assert!(self.pending_wt);
        self.pending_wt = false;
        let mut hold = std::mem::take(&mut self.hold);
        let dest = match &mut self.cbuf {
            Some(b) => b,
            None => &mut self.out,
        };
        dest.push_str(tag);
        dest.push_str(&hold);
        // keep hold's allocation for the next pending region
        hold.clear();
        self.hold = hold;
    }

    /// exact `<w:t>` seen in the stream
    fn on_wt(&mut self) {
        if self.pending_wt {
            // the previous `<w:t>`'s region ends here, un-triggered: plain
            self.flush_pending(WT_PLAIN);
        } else {
            // a previously flushed (retro) `<w:t>` is now decided plain
            self.retro_wt = None;
        }
        self.pending_wt = true;
    }

    /// `{{` or `{%` seen in the stream
    fn on_jinja_open(&mut self, two: &str) {
        // space_preserve: an opener before the next `<w:t>` upgrades the
        // pending (or, across a clean-region boundary, already-flushed) tag
        if self.pending_wt {
            self.flush_pending(WT_PRESERVE);
        } else if let Some(p) = self.retro_wt.take() {
            // insert the attribute before the tag's `>`
            self.out.insert_str(p + 4, " xml:space=\"preserve\"");
        }
        // clean_tags: an opener starts a region unless already inside one
        match &mut self.cbuf {
            Some(b) => b.push_str(two),
            None => {
                self.out.push_str(two);
                // reuse the previous region's buffer (capacity, no alloc)
                self.cbuf = Some(std::mem::take(&mut self.cbuf_spare));
            }
        }
    }

    /// `}}` or `%}` seen while inside a clean region
    fn on_clean_close(&mut self, two: &str) {
        let mut buf = self.cbuf.take().unwrap();
        if self.pending_wt {
            // The pending `<w:t>` sits inside this region but its
            // space_preserve decision is still open (a later opener, before
            // the next `<w:t>`, would still rewrite it in the sequential
            // composition). Flush it plain and remember where, so a later
            // opener can upgrade it in place. Cleaning the content before
            // and after the tag separately is equivalent to cleaning the
            // whole: the tag bytes are pure ASCII markup that none of
            // clean_tag_content's replacements can span or touch.
            self.pending_wt = false;
            clean_tag_content(&buf, &mut self.out);
            self.retro_wt = Some(self.out.len());
            self.out.push_str(WT_PLAIN);
            let mut hold = std::mem::take(&mut self.hold);
            clean_tag_content(&hold, &mut self.out);
            hold.clear();
            self.hold = hold;
        } else {
            clean_tag_content(&buf, &mut self.out);
        }
        // the closer is not consumed by clean_tags: it is plain outside text
        self.out.push_str(two);
        // keep the buffer for the next clean region
        buf.clear();
        self.cbuf_spare = buf;
    }

    fn on_token(&mut self, tok: Tok, text: &str) {
        match tok {
            Tok::Wt => self.on_wt(),
            Tok::Open => self.on_jinja_open(text),
            Tok::Close => self.on_clean_close(text),
        }
    }

    /// Feed a piece of post-strip text. Pieces may split tokens arbitrarily;
    /// undecided prefixes are kept in `carry` and resolved on the next feed.
    fn feed(&mut self, s: &str) {
        let mut rest = s;
        while !self.carry.is_empty() {
            match decide(&self.carry, self.cbuf.is_some()) {
                Dec::Undecided => {
                    let Some(&c) = rest.as_bytes().first() else {
                        return; // wait for the next feed / finish()
                    };
                    if c >= 0x80 {
                        // a non-ASCII byte can never complete a token: the
                        // carried bytes (pure ASCII) are plain text
                        let carried = std::mem::take(&mut self.carry);
                        self.push_plain(std::str::from_utf8(&carried).unwrap());
                    } else {
                        self.carry.push(c);
                        rest = &rest[1..]; // c is ASCII: still a char boundary
                    }
                }
                Dec::Plain => {
                    let first = self.carry.remove(0);
                    self.plain_dest().push(first as char); // carried bytes are ASCII
                }
                Dec::Token(tok, len) => {
                    let mut tmp = [0u8; 5];
                    tmp[..len].copy_from_slice(&self.carry[..len]);
                    self.carry.drain(..len);
                    let text = std::str::from_utf8(&tmp[..len]).unwrap();
                    self.on_token(tok, text);
                }
            }
        }
        if !rest.is_empty() {
            self.scan(rest);
        }
    }

    /// Scan one fed piece for token events, routing plain text through
    /// `push_plain`.
    fn scan(&mut self, s: &str) {
        let b = s.as_bytes();
        let n = b.len();
        let mut plain_start = 0usize;
        let mut i = 0usize;
        while i < n {
            // jump to the next event point: the exact `<w:t>` tag (memmem —
            // the hundreds of thousands of other `<...>` tags are skipped at
            // SIMD speed instead of one candidate check each) or a jinja
            // delimiter byte ('{', plus '}'/'%' inside a clean region)
            let hay = &b[i..];
            let p_wt = memchr::memmem::find(hay, b"<w:t>");
            let p_brace = if self.cbuf.is_some() {
                memchr::memchr3(b'{', b'}', b'%', hay)
            } else {
                memchr::memchr(b'{', hay)
            };
            let next = match (p_wt, p_brace) {
                (Some(a), Some(bb)) => i + a.min(bb),
                (Some(a), None) => i + a,
                (None, Some(bb)) => i + bb,
                (None, None) => break,
            };
            i = next;
            let c = b[i];
            let avail = n - i;
            let (tok, len) = match c {
                b'<' => (Some(Tok::Wt), 5), // memmem matched the exact tag
                b'{' if avail >= 2 => {
                    if b[i + 1] == b'{' || b[i + 1] == b'%' {
                        (Some(Tok::Open), 2)
                    } else {
                        (None, 1)
                    }
                }
                b'}' | b'%' if avail >= 2 => {
                    // only reached inside a clean region (see p_brace)
                    if b[i + 1] == b'}' {
                        (Some(Tok::Close), 2)
                    } else {
                        (None, 1)
                    }
                }
                _ => {
                    // possible token split at the end of this piece: stash it
                    self.push_plain(&s[plain_start..i]);
                    self.carry.extend_from_slice(&b[i..]);
                    return;
                }
            };
            match tok {
                None => i += 1,
                Some(t) => {
                    self.push_plain(&s[plain_start..i]);
                    self.on_token(t, &s[i..i + len]);
                    i += len;
                    plain_start = i;
                }
            }
        }
        // tail: the piece may end with a proper prefix of `<w:t>` (memmem
        // only surfaces full matches): stash it so the next feed can
        // complete the token
        let rem = &s[plain_start..];
        let rb = rem.as_bytes();
        let mut split = rb.len();
        for l in (1..=4usize.min(rb.len())).rev() {
            if rb[rb.len() - l..] == b"<w:t>"[..l] {
                split = rb.len() - l;
                break;
            }
        }
        self.push_plain(&rem[..split]);
        self.carry.extend_from_slice(&rb[split..]);
    }

    fn finish(mut self) -> String {
        if !self.carry.is_empty() {
            let carried = std::mem::take(&mut self.carry);
            self.push_plain(std::str::from_utf8(&carried).unwrap());
        }
        if self.pending_wt {
            self.flush_pending(WT_PLAIN); // EOF: no trigger can follow
        }
        if let Some(buf) = self.cbuf.take() {
            // clean_tags: a tag with no closer anywhere never matches; the
            // opener (already in out) and the content pass through verbatim
            self.out.push_str(&buf);
        }
        self.out
    }
}

/// Feed `s` to the emitter with every `</w:t>.*?(<w:t>|<w:t [^>]*>)` run
/// removed (same logic as strip_wt_pairs, but streaming).
fn feed_strip_wt_pairs(s: &str, em: &mut FusedEmitter) {
    if !s.contains("</w:t>") {
        em.feed(s);
        return;
    }
    let b = s.as_bytes();
    let mut copied = 0usize;
    let mut i = 0usize;
    while let Some(rel) = s[i..].find("</w:t>") {
        let close_start = i + rel;
        // earliest "<w:t>" or "<w:t ...>" opening tag after the close tag
        let mut j = close_start + 6;
        let open_end = loop {
            let Some(lt) = memchr::memchr(b'<', &b[j..]) else {
                break None;
            };
            j += lt;
            if j + 5 > b.len() {
                break None;
            }
            if b[j + 1..].starts_with(b"w:t") {
                if b[j + 4] == b'>' {
                    break Some(j + 5);
                }
                if b[j + 4] == b' ' {
                    match s[j + 5..].find('>') {
                        Some(g) => break Some(j + 5 + g + 1),
                        None => break None,
                    }
                }
            }
            j += 1;
        };
        match open_end {
            Some(e) => {
                em.feed(&s[copied..close_start]);
                copied = e;
                i = e;
            }
            None => break,
        }
    }
    em.feed(&s[copied..]);
}

/// Fused single-pass equivalent of running strip_tags_in_jinja_scan,
/// space_preserve_scan and clean_tags_scan in sequence (see FusedEmitter for
/// the equivalence argument). The strip driver is identical to
/// strip_tags_in_jinja_scan, except that the post-strip text is streamed
/// through the emitter instead of being materialized.
fn fused_jinja_scan(xml: &str) -> String {
    let mut em = FusedEmitter::new(xml.len());
    let mut rest = xml;
    loop {
        let b = rest.as_bytes();
        if b.len() < 2 {
            break;
        }
        // earliest opener among {% {# {{
        let mut o = None;
        let mut i = 0usize;
        while i + 1 < b.len() {
            if b[i] == b'{' && matches!(b[i + 1], b'{' | b'%' | b'#') {
                o = Some(i);
                break;
            }
            i += 1;
        }
        let Some(o) = o else { break };
        let closer = match b[o + 1] {
            b'%' => "%}",
            b'#' => "#}",
            _ => "}}",
        };
        let content_start = o + 2;
        let content_end = rest[content_start..]
            .find(closer)
            .map(|r| content_start + r)
            .unwrap_or(rest.len());
        em.feed(&rest[..content_start]);
        feed_strip_wt_pairs(&rest[content_start..content_end], &mut em);
        rest = &rest[content_end..];
    }
    em.feed(rest);
    em.finish()
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

    /// element_tag AC 驱动：多个 y 混合 + 非 ASCII 文本时，前一趟替换
    /// 移位后必须重新收集位置（回归：曾用失效位置切片导致 char boundary panic）
    #[test]
    fn test_element_tag_multi_y_cjk() {
        let cases: &[(&str, &str)] = &[
            (
                "<w:p><w:r><w:t>{%p 断 %}</w:t></w:r></w:p><w:tr><w:tc><w:p>{%tr for x in rows %}</w:p></w:tc></w:tr>",
                "{% 断 %}{% for x in rows %}",
            ),
            (
                "<w:tr><w:p>{%tr 断 %}</w:p></w:tr><w:p>{%p x %}</w:p>",
                "{% 断 %}{% x %}",
            ),
            (
                "<w:tc><w:p>{%tc 断 %}</w:p></w:tc><w:p>{%p y %}</w:p>",
                "{% 断 %}{% y %}",
            ),
            (
                "<w:p>{%p 断 %}</w:p><w:p>{%p 中文 %}</w:p><w:tr>{%tr z %}</w:tr>",
                "{% 断 %}{% 中文 %}{% z %}",
            ),
        ];
        for (input, expect) in cases {
            assert_eq!(patch_xml(input).as_ref(), *expect, "input: {}", input);
        }
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

    #[test]
    fn test_decode_text_entities_scan() {
        // 文本节点中的 &quot;/&apos;/数字引用被解码，标记字符保持转义
        assert_eq!(
            decode_text_entities("<w:t>a &quot;b&quot; &amp; &lt; &#8217; &#x4E2D;</w:t>"),
            "<w:t>a \"b\" &amp; &lt; ’ 中</w:t>"
        );
        // 只有 &amp;/&lt; 等保留实体时整体不变（借用，不拷贝）
        let s = "<w:t>a &amp; b &lt; c</w:t>";
        assert!(matches!(decode_text_entities(s), Cow::Borrowed(_)));
        assert_eq!(decode_text_entities(s), s);
        // 无 & 时恒等借用
        let s = "<w:t>plain</w:t>";
        assert!(matches!(decode_text_entities(s), Cow::Borrowed(_)));
        // 属性中的实体不动（只看 > < 之间的文本节点）
        let s = "<w:r w:attr=\"&quot;\"><w:t>&quot;</w:t></w:r>";
        assert_eq!(
            decode_text_entities(s),
            "<w:r w:attr=\"&quot;\"><w:t>\"</w:t></w:r>"
        );
        // 相邻标签 `><` 间无文本；`>` 在文本前的嵌套边界
        assert_eq!(decode_text_entities("<a><b>x</b></a>"), "<a><b>x</b></a>");
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

    #[test]
    fn test_strip_tags_in_jinja_scan_edges() {
        // 基本：剥掉标签内部跨 run 的 </w:t>...<w:t...>
        assert_eq!(
            strip_tags_in_jinja_scan("{{ x </w:t></w:r><w:r><w:t>+ y }}"),
            "{{ x + y }}"
        );
        // 带属性的 <w:t ...> 也剥
        assert_eq!(
            strip_tags_in_jinja_scan("{% a </w:t><w:t xml:space=\"preserve\">b %}"),
            "{% a b %}"
        );
        // 无闭合符时匹配到 EOF 且照常剥离（原正则不消耗闭合符）
        assert_eq!(
            strip_tags_in_jinja_scan("pre {{ x </w:t><w:t>y"),
            "pre {{ x y"
        );
        // 标签内的 </w:t> 后没有 <w:t> 开标签：不剥
        assert_eq!(
            strip_tags_in_jinja_scan("{{ a </w:t> b }}"),
            "{{ a </w:t> b }}"
        );
        // 闭合符不消耗，紧随其后的下一个标签照常处理
        assert_eq!(
            strip_tags_in_jinja_scan("{{ a }}</w:t><w:t>{{ b </w:t><w:t>c }}"),
            "{{ a }}</w:t><w:t>{{ b c }}"
        );
        // <w:t/> 不是开标签（正则只认 <w:t> 或 <w:t ...>）
        assert_eq!(
            strip_tags_in_jinja_scan("{{ a </w:t><w:t/>b }}"),
            "{{ a </w:t><w:t/>b }}"
        );
        // {# #} 注释同样处理
        assert_eq!(
            strip_tags_in_jinja_scan("{# n </w:t><w:t>x #}"),
            "{# n x #}"
        );
        // 恒等
        let s = "<w:p><w:r><w:t>plain</w:t></w:r></w:p>";
        assert_eq!(strip_tags_in_jinja_scan(s), s);
    }

    #[test]
    fn test_resolve_listing_scan() {
        // \n -> <w:br/>
        assert_eq!(
            resolve_listing("<w:p><w:r><w:t>a\nb</w:t></w:r></w:p>"),
            "<w:p><w:r><w:t>a</w:t><w:br/><w:t xml:space=\"preserve\">b</w:t></w:r></w:p>"
        );
        // \t -> tab run, run properties copied
        assert_eq!(
            resolve_listing("<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>a\tb</w:t></w:r></w:p>"),
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space=\"preserve\">b</w:t></w:r></w:p>"
        );
        // \x07 -> new paragraph, paragraph properties copied
        assert_eq!(
            resolve_listing("<w:p><w:pPr><w:jc/></w:pPr><w:r><w:t>a\u{7}b</w:t></w:r></w:p>"),
            "<w:p><w:pPr><w:jc/></w:pPr><w:r><w:t>a</w:t></w:r></w:p><w:p><w:pPr><w:jc/></w:pPr><w:r><w:t xml:space=\"preserve\">b</w:t></w:r></w:p>"
        );
        // \x0c -> page break paragraph
        assert_eq!(
            resolve_listing("<w:p><w:r><w:t>a\u{c}b</w:t></w:r></w:p>"),
            "<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:br w:type=\"page\"/></w:r></w:p><w:p><w:r><w:t xml:space=\"preserve\">b</w:t></w:r></w:p>"
        );
        // w:t with attributes; paragraph with attributes
        assert_eq!(
            resolve_listing("<w:p x=\"1\"><w:r><w:t xml:space=\"preserve\">a\nb</w:t></w:r></w:p>"),
            "<w:p x=\"1\"><w:r><w:t xml:space=\"preserve\">a</w:t><w:br/><w:t xml:space=\"preserve\">b</w:t></w:r></w:p>"
        );
        // identity: no listing chars
        let s = "<w:p><w:r><w:t>plain</w:t></w:r></w:p>";
        assert!(matches!(resolve_listing(s), Cow::Borrowed(_)));
        // <w:pPr>/<w:rPr>/<w:tab> are not paragraph/run/text openings
        assert_eq!(
            resolve_listing("<w:pPr><w:jc/></w:pPr><w:tab/>\n"),
            "<w:pPr><w:jc/></w:pPr><w:tab/>\n"
        );
        // paragraph without close tag: verbatim
        assert_eq!(resolve_listing("<w:p><w:r><w:t>a\nb"), "<w:p><w:r><w:t>a\nb");
        // listing char outside <w:t> (but inside a paragraph) is untouched
        assert_eq!(
            resolve_listing("<w:p>\n<w:r><w:t>x</w:t></w:r></w:p>"),
            "<w:p>\n<w:r><w:t>x</w:t></w:r></w:p>"
        );
    }

    #[test]
    fn test_clean_tags_scan_edges() {
        // 闭合符取第一个 }} 或 %}，与开符类型无关
        assert_eq!(clean_tags_scan("{{ a &lt; b %}"), "{{ a < b %}");
        assert_eq!(clean_tags_scan("{% a &gt; b }}"), "{% a > b }}");
        // 空标签
        assert_eq!(clean_tags_scan("{{}}"), "{{}}");
        assert_eq!(clean_tags_scan("{%%}"), "{%%}");
        // 无闭合符：整体原样
        assert_eq!(clean_tags_scan("{{ a &lt; b"), "{{ a &lt; b");
        // 标签外的实体不动
        assert_eq!(
            clean_tags_scan("x &lt; {{ a &lt; b }} y &lt;"),
            "x &lt; {{ a < b }} y &lt;"
        );
        // &#8216; 与智能引号
        assert_eq!(clean_tags_scan("{{ &#8216;x&#8216; }}"), "{{ 'x' }}");
        assert_eq!(
            clean_tags_scan("{{ \u{2018}a\u{2019}\u{201c}b\u{201d} }}"),
            "{{ 'a'\"b\" }}"
        );
        // 其他 E2 80 xx 字符（如 U+201A）不动
        assert_eq!(clean_tags_scan("{{ \u{201a} }}"), "{{ \u{201a} }}");
        // 恒等
        let s = "<w:p><w:r><w:t>plain</w:t></w:r></w:p>";
        assert_eq!(clean_tags_scan(s), s);
    }

    // ------------------------------------------------------------------
    // fused_jinja_scan 等价性：单趟融合扫描必须产生与按序执行三个旧 pass
    // 完全一致的输出（gate 语义与 patch_xml 一致，逐 pass 条件应用）
    // ------------------------------------------------------------------

    /// 参考实现：按 patch_xml 的 gate 逻辑逐 pass 条件应用三个旧扫描函数
    fn ref_three_passes(xml: &str) -> String {
        let mut s = xml.to_string();
        if contains_any(&s, &["{{", "{%", "{#"]) {
            s = strip_tags_in_jinja_scan(&s);
        }
        if s.contains("<w:t>") && contains_any(&s, &["{{", "{%"]) {
            s = space_preserve_scan(&s);
        }
        if contains_any(&s, &["{{", "{%"]) {
            s = clean_tags_scan(&s);
        }
        s
    }

    fn assert_fused_equiv(xml: &str) {
        assert_eq!(fused_jinja_scan(xml), ref_three_passes(xml), "input: {:?}", xml);
    }

    /// 现有 strip/space/clean 单元测试与 patch_xml 端到端测试的全部输入
    #[test]
    fn test_fused_jinja_equiv_existing_inputs() {
        let inputs = [
            // strip_tags_in_jinja_scan 测试输入
            "{{ x </w:t></w:r><w:r><w:t>+ y }}",
            "{% a </w:t><w:t xml:space=\"preserve\">b %}",
            "pre {{ x </w:t><w:t>y",
            "{{ a </w:t> b }}",
            "{{ a }}</w:t><w:t>{{ b </w:t><w:t>c }}",
            "{{ a </w:t><w:t/>b }}",
            "{# n </w:t><w:t>x #}",
            // space_preserve_scan 测试输入
            "<w:t>{{ x }}</w:t>",
            "<w:t>{% if x %}</w:t>",
            "<w:t>a</w:t><w:t>{{b}}</w:t>",
            "<w:t>a</w:t>{{ x }}<w:t>b</w:t>",
            "<w:t xml:space=\"preserve\">{{x}}</w:t>",
            // clean_tags_scan 测试输入
            "{{ a &lt; b %}",
            "{% a &gt; b }}",
            "{{}}",
            "{%%}",
            "{{ a &lt; b",
            "x &lt; {{ a &lt; b }} y &lt;",
            "{{ &#8216;x&#8216; }}",
            "{{ \u{2018}a\u{2019}\u{201c}b\u{201d} }}",
            "{{ \u{201a} }}",
            // patch_xml 端到端测试输入
            "<w:p><w:r><w:t>{{ x </w:t></w:r><w:r><w:t>+ y }}</w:t></w:r></w:p>",
            "<w:p><w:r><w:t>{</w:t></w:r><w:r><w:t>{ x }</w:t></w:r></w:p>",
            "<w:tc><w:tcPr><w:gridSpan w:val=\"9\"/></w:tcPr><w:r><w:t>{% colspan span %}</w:t></w:r></w:tc>",
            "<w:tc><w:tcPr><w:shd w:val=\"clear\" w:fill=\"FF0000\"/></w:tcPr><w:r><w:t>{% cellbg color %}</w:t></w:r></w:tc>",
            "<w:body><w:p><w:r><w:t>{{r rt }}</w:t></w:r></w:p></w:body>",
            "<w:p><w:r><w:t>text</w:t></w:r></w:p><w:p><w:r><w:t>{%- if x %}</w:t></w:r></w:p>",
            "<w:p><w:r><w:t>{% endif -%}</w:t></w:r></w:p><w:p><w:r><w:t>next</w:t></w:r></w:p>",
            "<w:p><w:r><w:t>{%p if x %}</w:t></w:r></w:p>",
            "<w:tr><w:tc><w:p><w:r><w:t>{{tr row }}</w:t></w:r></w:p></w:tc></w:tr>",
            "<w:p><w:r><w:t>{#p note #}</w:t></w:r></w:p>",
            "<w:tc><w:tcPr></w:tcPr><w:r><w:t>{% vm %}content</w:t></w:r></w:tc>",
            "<w:tc><w:tcPr></w:tcPr><w:r><w:t>a{% hm %}b</w:t></w:r></w:tc>",
            "<w:tc><w:tcPr><w:gridSpan w:val=\"3\"/></w:tcPr><w:r><w:t>{% hm %}</w:t></w:r></w:tc>",
            "<w:p><w:r><w:t>{{ x &lt; y }}</w:t></w:r></w:p>",
            "<w:p><w:r><w:t>{{ \u{201c}s\u{201d} }}</w:t></w:r></w:p>",
            // gate 不满足：三个旧 pass 都不会被调用，融合扫描同样恒等
            "<w:p><w:r><w:t>plain</w:t></w:r></w:p>",
            "<w:t>no markers here</w:t>",
            "",
        ];
        for x in inputs {
            assert_fused_equiv(x);
        }
    }

    /// 合成边界情况：跨 run 拆分、无闭符、闭符规则分歧、{# 内嵌 {{、
    /// jinja 内容里孤立的 <w:t>（retro 升级路径）、carry 跨 feed 边界等
    #[test]
    fn test_fused_jinja_equiv_edge_cases() {
        let inputs = [
            // clean 开符不含 {#，但 {# 区域内的 {{ 子开符会被 clean 处理
            "{# {{ a &lt; b }} #}",
            "{# a {{ b }} c #}",
            "{# a {{ b #} c }}",
            "{{ a {# b }} c #}",
            // clean 闭符（第一个 }}/%}）先于 strip 闭符（按开符类型）
            "{% a }} b %}",
            "{{ a %} b }}",
            "{% a &lt; }} b &gt; %}",
            // jinja 内容里孤立的 <w:t>（无前置 </w:t>，strip 不删）：
            // 它是 clean 区域内的 space_preserve 候选
            "{{ a <w:t> b }}",
            "{{ a <w:t> {% b %} c }}",
            // retro 路径：<w:t> 在 clean 区域内，触发它的 {% 在区域关闭后
            "{{ a <w:t> b }} {% c %}",
            "<w:t>{{ a <w:t> b }}</w:t>",
            "{{ <w:t> x &lt; y }} z {% w %}",
            // 多个 retro/触发交错
            "{{ <w:t> }}t{% u %}{{ <w:t> }}v{% w %}",
            // 跨 run 拆分标签 + space/clean 组合
            "<w:p><w:r><w:t>{{ x </w:t></w:r><w:r><w:t>&lt; y }}</w:t></w:r></w:p>",
            "<w:t>{{ a </w:t></w:r><w:r><w:t>\u{201c}b\u{201d} }}</w:t>",
            // strip 拼接出新的 {{（旧 strip 看不到，但 space/clean 能看到）
            "{{ a {</w:t><w:t>{ b }}",
            "{{ x </w:t><w:t>{ }}",
            // {%- / -%} 对这三个 pass 无特殊语义，照常处理
            "{%- if x %}",
            "<w:t>{%- x -%}</w:t>",
            // 无闭符：strip 到 EOF 照常剥离；clean 无闭符则整体原样
            "pre {{ x </w:t><w:t>y",
            "{{ a &lt; b",
            "<w:t>{{ a </w:t><w:t> &lt; b",
            // 紧邻多个标签、空标签
            "{{ a }}{% b %}{# c #}",
            "{{}}{%%}{##}",
            // 嵌套引号与智能引号、&#8216;
            "{{ \"a\" &#8216;b&#8216; \u{2018}c\u{2019} }}",
            "{% if a == \u{201c}x\u{201d} and b &lt; 2 %}",
            // 多字节字符环绕 token（carry 不得切开 UTF-8 字符）
            "<w:t>中文{{ x }}日本語</w:t>",
            "{{ x </w:t><w:t>é }}",
            // carry 边界：feed 片段以 { 或 <w:t 结尾（由 strip 切分产生）
            "{{ a {</w:t><w:t>{ b }}",
            "{{ a </w:t><w:t",
            "{% k </w:t><w:t",
            // EOF 时 pending <w:t> 未触发：保持原样
            "<w:t>{{ x }}</w:t><w:t>tail",
            "<w:t>plain</w:t><w:t>{{ y }}",
            // 连续 <w:t>：前一个区域无触发
            "<w:t>a<w:t>{{ b }}",
            // closer 紧贴 opener
            "<w:t>{{}}</w:t>",
            "<w:t>{% x %}{% y %}</w:t>",
        ];
        for x in inputs {
            assert_fused_equiv(x);
        }
    }

    /// 长文档等价性（同时作为性能冒烟测试）：大量 run + 周期性跨 run
    /// 拆分的 jinja 标签
    #[test]
    fn test_fused_jinja_equiv_long() {
        let mut long = String::new();
        for i in 0..5_000 {
            if i % 7 == 3 {
                long.push_str("<w:p><w:r><w:t>{{ x </w:t></w:r><w:r><w:t>+ &lt; y }}</w:t></w:r></w:p>");
            } else {
                long.push_str("<w:p><w:r><w:t>t</w:t></w:r></w:p>");
            }
        }
        assert_fused_equiv(&long);
    }
}



