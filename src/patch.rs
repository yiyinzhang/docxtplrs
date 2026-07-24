//! Port of docxtpl's patch_xml / resolve_listing logic (regex-based XML cleaning).

use fancy_regex::{Captures, Regex};

fn re(pattern: &str) -> Regex {
    if std::env::var("PATCH_DEBUG").is_ok() {
        eprintln!("[patch] running: {}", pattern);
    }
    fancy_regex::RegexBuilder::new(pattern)
        .backtrack_limit(50_000_000)
        .build()
        .unwrap_or_else(|e| panic!("invalid regex {}: {}", pattern, e))
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
    let mut xml = sub_str(
        r"(?s)(?<=\{)(?>(?:<[^>]*>)+)(?=[\{%\#])|(?<=[%\}\#])(?>(?:<[^>]*>)+)(?=\})",
        "",
        src_xml,
    );

    // replace {{<some tags>jinja2 stuff<some other tags>}} by {{jinja2 stuff}}
    xml = sub(
        r"(?s)\{%(?:(?!%\}).)*|\{#(?:(?!#\}).)*|\{\{(?:(?!}\}).)*",
        |m| sub_str(r"(?s)</w:t>.*?(<w:t>|<w:t [^>]*>)", "", m.get(0).unwrap().as_str()),
        &xml,
    );

    // manage table cell colspan
    xml = sub(
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
    );

    // manage table cell background color
    xml = sub(
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
    );

    // ensure space preservation (hand-rolled linear scan; the original
    // tempered-dot regex overflows the backtrack stack on large documents)
    xml = space_preserve_scan(&xml);
    xml = sub(
        r"(?s)(\{\{r\s.*?\}\}|\{%r\s.*?\%\})",
        |m| {
            format!(
                "</w:t></w:r><w:r><w:t xml:space=\"preserve\">{}</w:t></w:r><w:r><w:t xml:space=\"preserve\">",
                m.get(1).unwrap().as_str()
            )
        },
        &xml,
    );

    // {%- will merge with previous paragraph text (hand-rolled; the
    // original regex spans paragraphs and overflows the backtrack stack)
    xml = dash_merge_prev_scan(&xml);
    // -%} will merge with next paragraph text
    xml = dash_merge_next_scan(&xml);

    // replace into xml code the row/paragraph/run containing
    // {%y xxx %} / {{y xxx}} / {#y xxx #} template tag by the tag alone
    // without any surrounding <w:y> tags (hand-rolled linear scan; the
    // original tempered-dot regex overflows the backtrack stack on large
    // documents)
    for y in ["tr", "tc", "p", "r"] {
        xml = element_tag_scan(&xml, y, false);
    }
    for y in ["tr", "tc", "p"] {
        xml = element_tag_scan(&xml, y, true);
    }

    // add vMerge
    // use {% vm %} to make this table cell and its copies
    // be vertically merged within a {% for %}
    xml = sub(
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
    );

    // Use {% hm %} to make table cell become horizontally merged within a {% for %}.
    xml = sub(
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
    );

    // clean tags: unescape entities and smart quotes inside jinja tags
    xml = sub(r"(?<=\{[\{%])(.*?)(?=[\}%]\})", |m| {
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
    }, &xml);

    xml
}

/// replace only first occurrence
fn sub_first(pattern: &str, replacement: &str, text: &str) -> String {
    let rex = re(pattern);
    rex.replace(text, replacement).to_string()
}

/// Port of DocxTemplate.resolve_listing
pub fn resolve_listing(xml: &str) -> String {
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

    sub(
        r"(?s)<w:p(?: [^>]*)?>.*?</w:p>",
        |m| resolve_paragraph(m),
        xml,
    )
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
    let markers: Vec<(&str, &str)> = if comment {
        vec![(Box::leak(format!("{{#{} ", y).into_boxed_str()), "#}")]
    } else {
        vec![
            (Box::leak(format!("{{{{{} ", y).into_boxed_str()), "}}"),
            (Box::leak(format!("{{%{} ", y).into_boxed_str()), "%}"),
        ]
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
