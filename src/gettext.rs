//! GNU gettext .mo catalog parsing and plural-form evaluation,
//! powering `{% trans %}` support.

use std::collections::HashMap;

#[derive(Debug, Default)]
pub struct Catalog {
    /// msgid (singular) -> plural forms (index 0 = singular form)
    pub messages: HashMap<String, Vec<String>>,
    /// context\x04msgid -> forms (pgettext)
    pub contextual: HashMap<String, Vec<String>>,
    pub nplurals: usize,
    pub plural_rule: Option<PluralExpr>,
}

fn le_u32(b: &[u8], off: usize) -> u32 {
    u32::from_le_bytes([b[off], b[off + 1], b[off + 2], b[off + 3]])
}

impl Catalog {
    /// Parse a .mo file (little-endian GNU format).
    pub fn parse(data: &[u8]) -> Result<Catalog, String> {
        if data.len() < 28 || le_u32(data, 0) != 0x950412de {
            return Err("not a valid .mo file".to_string());
        }
        let n = le_u32(data, 8) as usize;
        let orig_off = le_u32(data, 12) as usize;
        let trans_off = le_u32(data, 16) as usize;

        let mut catalog = Catalog {
            nplurals: 2,
            ..Default::default()
        };

        let read_str = |table: usize, i: usize| -> String {
            let len = le_u32(data, table + i * 8) as usize;
            let off = le_u32(data, table + i * 8 + 4) as usize;
            if off + len > data.len() {
                return String::new();
            }
            String::from_utf8_lossy(&data[off..off + len]).to_string()
        };

        for i in 0..n {
            let key = read_str(orig_off, i);
            let val = read_str(trans_off, i);
            if key.is_empty() {
                catalog.parse_header(&val);
                continue;
            }
            let forms: Vec<String> = val.split('\0').map(|s| s.to_string()).collect();
            let mut parts = key.splitn(2, '\0');
            let singular = parts.next().unwrap_or("").to_string();
            if let Some((ctx, msgid)) = singular.split_once('\u{4}') {
                catalog
                    .contextual
                    .insert(format!("{}\u{4}{}", ctx, msgid), forms);
            } else {
                catalog.messages.insert(singular, forms);
            }
        }
        Ok(catalog)
    }

    fn parse_header(&mut self, header: &str) {
        for line in header.split('\n') {
            if let Some(rest) = line.strip_prefix("Plural-Forms:") {
                let rest = rest.trim();
                if let Some(n) = rest.strip_prefix("nplurals=") {
                    let num: String = n.chars().take_while(|c| c.is_ascii_digit()).collect();
                    self.nplurals = num.parse().unwrap_or(2);
                }
                if let Some(p) = rest.find("plural=") {
                    let expr = rest[p + 7..].trim_end_matches(';').trim().to_string();
                    if let Some(parsed) = PluralExpr::parse(&expr) {
                        self.plural_rule = Some(parsed);
                    }
                }
            }
        }
    }

    pub fn gettext(&self, msgid: &str) -> String {
        self.messages
            .get(msgid)
            .and_then(|f| f.first().cloned())
            .unwrap_or_else(|| msgid.to_string())
    }

    pub fn pgettext(&self, context: &str, msgid: &str) -> String {
        self.contextual
            .get(&format!("{}\u{4}{}", context, msgid))
            .and_then(|f| f.first().cloned())
            .unwrap_or_else(|| msgid.to_string())
    }

    pub fn ngettext(&self, singular: &str, plural: &str, n: i64) -> String {
        if let Some(forms) = self.messages.get(singular) {
            let idx = self.plural_index(n).min(forms.len().saturating_sub(1));
            if let Some(form) = forms.get(idx) {
                return form.clone();
            }
        }
        if n == 1 {
            singular.to_string()
        } else {
            plural.to_string()
        }
    }

    pub fn npgettext(&self, context: &str, singular: &str, plural: &str, n: i64) -> String {
        if let Some(forms) = self.contextual.get(&format!("{}\u{4}{}", context, singular)) {
            let idx = self.plural_index(n).min(forms.len().saturating_sub(1));
            if let Some(form) = forms.get(idx) {
                return form.clone();
            }
        }
        if n == 1 {
            singular.to_string()
        } else {
            plural.to_string()
        }
    }

    fn plural_index(&self, n: i64) -> usize {
        match &self.plural_rule {
            Some(rule) => {
                let v = rule.eval(n);
                (if v != 0 { v } else { 0 }) as usize % self.nplurals.max(1)
            }
            None => {
                if n != 1 {
                    1
                } else {
                    0
                }
            }
        }
    }
}

// ---------------- plural expression parser ----------------

#[derive(Debug, Clone)]
pub enum PluralExpr {
    Num(i64),
    N,
    Ternary(Box<PluralExpr>, Box<PluralExpr>, Box<PluralExpr>),
    Or(Box<PluralExpr>, Box<PluralExpr>),
    And(Box<PluralExpr>, Box<PluralExpr>),
    Eq(Box<PluralExpr>, Box<PluralExpr>),
    Ne(Box<PluralExpr>, Box<PluralExpr>),
    Lt(Box<PluralExpr>, Box<PluralExpr>),
    Le(Box<PluralExpr>, Box<PluralExpr>),
    Gt(Box<PluralExpr>, Box<PluralExpr>),
    Ge(Box<PluralExpr>, Box<PluralExpr>),
    Mod(Box<PluralExpr>, Box<PluralExpr>),
    Add(Box<PluralExpr>, Box<PluralExpr>),
    Sub(Box<PluralExpr>, Box<PluralExpr>),
    Not(Box<PluralExpr>),
}

impl PluralExpr {
    pub fn parse(s: &str) -> Option<PluralExpr> {
        let mut p = PluralParser {
            chars: s.chars().collect(),
            pos: 0,
        };
        let e = p.parse_ternary()?;
        p.skip_ws();
        if p.pos == p.chars.len() {
            Some(e)
        } else {
            None
        }
    }

    pub fn eval(&self, n: i64) -> i64 {
        use PluralExpr::*;
        match self {
            Num(v) => *v,
            N => n,
            Ternary(c, a, b) => {
                if c.eval(n) != 0 {
                    a.eval(n)
                } else {
                    b.eval(n)
                }
            }
            Or(a, b) => ((a.eval(n) != 0) || (b.eval(n) != 0)) as i64,
            And(a, b) => ((a.eval(n) != 0) && (b.eval(n) != 0)) as i64,
            Eq(a, b) => (a.eval(n) == b.eval(n)) as i64,
            Ne(a, b) => (a.eval(n) != b.eval(n)) as i64,
            Lt(a, b) => (a.eval(n) < b.eval(n)) as i64,
            Le(a, b) => (a.eval(n) <= b.eval(n)) as i64,
            Gt(a, b) => (a.eval(n) > b.eval(n)) as i64,
            Ge(a, b) => (a.eval(n) >= b.eval(n)) as i64,
            Mod(a, b) => {
                let d = b.eval(n);
                if d == 0 {
                    0
                } else {
                    a.eval(n) % d
                }
            }
            Add(a, b) => a.eval(n) + b.eval(n),
            Sub(a, b) => a.eval(n) - b.eval(n),
            Not(a) => (a.eval(n) == 0) as i64,
        }
    }
}

struct PluralParser {
    chars: Vec<char>,
    pos: usize,
}

impl PluralParser {
    fn peek(&self) -> Option<char> {
        self.chars.get(self.pos).copied()
    }
    fn next(&mut self) -> Option<char> {
        let c = self.peek();
        if c.is_some() {
            self.pos += 1;
        }
        c
    }
    fn skip_ws(&mut self) {
        while self.peek().map(|c| c.is_whitespace()).unwrap_or(false) {
            self.pos += 1;
        }
    }
    fn expect(&mut self, s: &str) -> bool {
        self.skip_ws();
        for c in s.chars() {
            if self.next() != Some(c) {
                return false;
            }
        }
        true
    }

    fn parse_ternary(&mut self) -> Option<PluralExpr> {
        let cond = self.parse_or()?;
        self.skip_ws();
        if self.peek() == Some('?') {
            self.next();
            let a = self.parse_ternary()?;
            if !self.expect(":") {
                return None;
            }
            let b = self.parse_ternary()?;
            Some(PluralExpr::Ternary(Box::new(cond), Box::new(a), Box::new(b)))
        } else {
            Some(cond)
        }
    }

    fn parse_or(&mut self) -> Option<PluralExpr> {
        let mut left = self.parse_and()?;
        loop {
            self.skip_ws();
            if self.peek() == Some('|') {
                if !self.expect("||") {
                    return None;
                }
                let right = self.parse_and()?;
                left = PluralExpr::Or(Box::new(left), Box::new(right));
            } else {
                return Some(left);
            }
        }
    }

    fn parse_and(&mut self) -> Option<PluralExpr> {
        let mut left = self.parse_cmp()?;
        loop {
            self.skip_ws();
            if self.peek() == Some('&') {
                if !self.expect("&&") {
                    return None;
                }
                let right = self.parse_cmp()?;
                left = PluralExpr::And(Box::new(left), Box::new(right));
            } else {
                return Some(left);
            }
        }
    }

    fn parse_cmp(&mut self) -> Option<PluralExpr> {
        let left = self.parse_additive()?;
        self.skip_ws();
        let save = self.pos;
        for (op, make) in [
            ("==", PluralExpr::Eq as fn(Box<PluralExpr>, Box<PluralExpr>) -> PluralExpr),
            ("!=", PluralExpr::Ne),
            ("<=", PluralExpr::Le),
            (">=", PluralExpr::Ge),
            ("<", PluralExpr::Lt),
            (">", PluralExpr::Gt),
        ] {
            self.pos = save;
            if self.expect(op) {
                let right = self.parse_additive()?;
                return Some(make(Box::new(left), Box::new(right)));
            }
        }
        self.pos = save;
        Some(left)
    }

    fn parse_additive(&mut self) -> Option<PluralExpr> {
        let mut left = self.parse_multiplicative()?;
        loop {
            self.skip_ws();
            match self.peek() {
                Some('+') => {
                    self.next();
                    let right = self.parse_multiplicative()?;
                    left = PluralExpr::Add(Box::new(left), Box::new(right));
                }
                Some('-') => {
                    self.next();
                    let right = self.parse_multiplicative()?;
                    left = PluralExpr::Sub(Box::new(left), Box::new(right));
                }
                _ => return Some(left),
            }
        }
    }

    // C/gettext precedence: `%` is multiplicative, binding tighter than `+`/`-`.
    fn parse_multiplicative(&mut self) -> Option<PluralExpr> {
        let mut left = self.parse_unary()?;
        loop {
            self.skip_ws();
            if self.peek() == Some('%') {
                self.next();
                let right = self.parse_unary()?;
                left = PluralExpr::Mod(Box::new(left), Box::new(right));
            } else {
                return Some(left);
            }
        }
    }

    fn parse_unary(&mut self) -> Option<PluralExpr> {
        self.skip_ws();
        match self.peek() {
            Some('!') => {
                self.next();
                Some(PluralExpr::Not(Box::new(self.parse_unary()?)))
            }
            Some('(') => {
                self.next();
                let e = self.parse_ternary()?;
                if !self.expect(")") {
                    return None;
                }
                Some(e)
            }
            Some('n') => {
                self.next();
                Some(PluralExpr::N)
            }
            Some(c) if c.is_ascii_digit() => {
                let mut num = String::new();
                while let Some(c) = self.peek() {
                    if c.is_ascii_digit() {
                        num.push(c);
                        self.next();
                    } else {
                        break;
                    }
                }
                Some(PluralExpr::Num(num.parse().ok()?))
            }
            _ => None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_plural_rules() {
        let e = PluralExpr::parse("n != 1").unwrap();
        assert_eq!(e.eval(1), 0);
        assert_eq!(e.eval(2), 1);
        let e = PluralExpr::parse("(n%10==1 && n%100!=11) ? 0 : 1").unwrap();
        assert_eq!(e.eval(1), 0);
        assert_eq!(e.eval(11), 1);
        assert_eq!(e.eval(21), 0);
    }

    // ---------------- .mo builder helper ----------------

    /// Build a minimal little-endian GNU .mo byte stream:
    /// 28-byte header (magic, version=0, N, orig/trans table offsets, no hash
    /// table), two tables of N (len, offset) pairs, then raw string bytes.
    fn build_mo(entries: &[(&[u8], &[u8])]) -> Vec<u8> {
        let n = entries.len();
        let orig_off = 28usize;
        let trans_off = orig_off + n * 8;
        let mut data_off = trans_off + n * 8;
        let mut orig_table = Vec::new();
        let mut trans_table = Vec::new();
        let mut strings = Vec::new();
        for (key, val) in entries {
            orig_table.extend_from_slice(&(key.len() as u32).to_le_bytes());
            orig_table.extend_from_slice(&(data_off as u32).to_le_bytes());
            strings.extend_from_slice(key);
            data_off += key.len();
            trans_table.extend_from_slice(&(val.len() as u32).to_le_bytes());
            trans_table.extend_from_slice(&(data_off as u32).to_le_bytes());
            strings.extend_from_slice(val);
            data_off += val.len();
        }
        let mut mo = Vec::new();
        mo.extend_from_slice(&0x950412deu32.to_le_bytes()); // magic
        mo.extend_from_slice(&0u32.to_le_bytes()); // version
        mo.extend_from_slice(&(n as u32).to_le_bytes());
        mo.extend_from_slice(&(orig_off as u32).to_le_bytes());
        mo.extend_from_slice(&(trans_off as u32).to_le_bytes());
        mo.extend_from_slice(&0u32.to_le_bytes()); // hash table size
        mo.extend_from_slice(&0u32.to_le_bytes()); // hash table offset
        mo.extend_from_slice(&orig_table);
        mo.extend_from_slice(&trans_table);
        mo.extend_from_slice(&strings);
        mo
    }

    const HEADER_2PLURAL: &[u8] =
        b"Project-Id-Version: t\nPlural-Forms: nplurals=2; plural=n!=1;\n";

    // ---------------- Catalog::parse ----------------

    #[test]
    fn test_parse_mo_rejects_truncated_data() {
        assert!(Catalog::parse(&[]).is_err());
        assert!(Catalog::parse(&[0xde, 0x12, 0x04]).is_err());
        // correct magic but shorter than the 28-byte header
        assert!(Catalog::parse(&0x950412deu32.to_le_bytes()).is_err());
    }

    #[test]
    fn test_parse_mo_rejects_bad_magic() {
        let mut mo = build_mo(&[]);
        mo[0] ^= 0xff; // corrupt magic, keep length >= 28
        assert_eq!(Catalog::parse(&mo).unwrap_err(), "not a valid .mo file");
    }

    #[test]
    fn test_parse_mo_empty_catalog() {
        let cat = Catalog::parse(&build_mo(&[])).unwrap();
        assert!(cat.messages.is_empty());
        assert!(cat.contextual.is_empty());
        assert_eq!(cat.nplurals, 2); // documented default
        assert!(cat.plural_rule.is_none());
    }

    #[test]
    fn test_parse_mo_simple_mapping() {
        let mo = build_mo(&[(b"", HEADER_2PLURAL), (b"hello", b"bonjour")]);
        let cat = Catalog::parse(&mo).unwrap();
        assert_eq!(cat.messages["hello"], vec!["bonjour".to_string()]);
    }

    #[test]
    fn test_parse_mo_header_plural_forms() {
        let mo = build_mo(&[(
            b"",
            b"Plural-Forms: nplurals=3; plural=n==1 ? 0 : n==2 ? 1 : 2;\n",
        )]);
        let cat = Catalog::parse(&mo).unwrap();
        assert_eq!(cat.nplurals, 3);
        let rule = cat.plural_rule.expect("plural rule should parse");
        assert_eq!(rule.eval(1), 0);
        assert_eq!(rule.eval(2), 1);
        assert_eq!(rule.eval(5), 2);
    }

    #[test]
    fn test_parse_mo_header_without_plural_forms() {
        let mo = build_mo(&[(b"", b"Content-Type: text/plain; charset=UTF-8\n")]);
        let cat = Catalog::parse(&mo).unwrap();
        assert_eq!(cat.nplurals, 2);
        assert!(cat.plural_rule.is_none());
    }

    #[test]
    fn test_parse_mo_header_unparseable_plural_expr_keeps_none() {
        let mo = build_mo(&[(b"", b"Plural-Forms: nplurals=4; plural=n***;\n")]);
        let cat = Catalog::parse(&mo).unwrap();
        assert_eq!(cat.nplurals, 4); // nplurals still picked up
        assert!(cat.plural_rule.is_none()); // bad expression ignored
    }

    #[test]
    fn test_parse_mo_plural_forms_split_on_nul() {
        let mo = build_mo(&[
            (b"", HEADER_2PLURAL),
            (b"apple\0apples", b"pomme\0pommes"),
        ]);
        let cat = Catalog::parse(&mo).unwrap();
        // key is truncated at the first NUL; value keeps both forms
        assert_eq!(
            cat.messages["apple"],
            vec!["pomme".to_string(), "pommes".to_string()]
        );
    }

    #[test]
    fn test_parse_mo_contextual_entry() {
        let mo = build_mo(&[(b"", HEADER_2PLURAL), (b"menu\x04file", b"fichier")]);
        let cat = Catalog::parse(&mo).unwrap();
        assert!(cat.messages.is_empty());
        assert_eq!(cat.contextual["menu\x04file"], vec!["fichier".to_string()]);
    }

    #[test]
    fn test_parse_mo_out_of_bounds_strings_become_empty() {
        // one entry whose string data points past EOF; the table itself is in bounds
        let mut mo = build_mo(&[(b"k", b"v")]);
        let len = mo.len();
        // corrupt the orig string offset (first table entry, bytes 32..36) to far past EOF
        mo[32..36].copy_from_slice(&10_000u32.to_le_bytes());
        let cat = Catalog::parse(&mo).unwrap();
        // key reads as empty -> treated as the header entry, so no message is stored
        assert!(cat.messages.is_empty());
        let _ = len;
    }

    #[test]
    fn test_parse_mo_invalid_utf8_is_lossy() {
        let mo = build_mo(&[(b"\xff\xfe", b"ok")]);
        let cat = Catalog::parse(&mo).unwrap();
        assert_eq!(cat.messages["\u{fffd}\u{fffd}"], vec!["ok".to_string()]);
    }

    // ---------------- lookup / fallback ----------------

    #[test]
    fn test_gettext_returns_translation() {
        let mo = build_mo(&[(b"", HEADER_2PLURAL), (b"hello", b"bonjour")]);
        let cat = Catalog::parse(&mo).unwrap();
        assert_eq!(cat.gettext("hello"), "bonjour");
    }

    #[test]
    fn test_gettext_missing_msgid_returns_msgid() {
        let mo = build_mo(&[(b"", HEADER_2PLURAL), (b"hello", b"bonjour")]);
        let cat = Catalog::parse(&mo).unwrap();
        assert_eq!(cat.gettext("goodbye"), "goodbye");
        // empty (default) catalog: everything is identity
        assert_eq!(Catalog::default().gettext("hello"), "hello");
    }

    #[test]
    fn test_pgettext_returns_contextual_translation() {
        let mo = build_mo(&[(b"", HEADER_2PLURAL), (b"menu\x04file", b"fichier")]);
        let cat = Catalog::parse(&mo).unwrap();
        assert_eq!(cat.pgettext("menu", "file"), "fichier");
    }

    #[test]
    fn test_pgettext_missing_returns_msgid() {
        let mo = build_mo(&[(b"", HEADER_2PLURAL), (b"menu\x04file", b"fichier")]);
        let cat = Catalog::parse(&mo).unwrap();
        // wrong context or wrong msgid falls back to the untranslated msgid
        assert_eq!(cat.pgettext("edit", "file"), "file");
        assert_eq!(cat.pgettext("menu", "edit"), "edit");
        assert_eq!(Catalog::default().pgettext("menu", "file"), "file");
    }

    #[test]
    fn test_ngettext_selects_form_by_plural_rule() {
        let mo = build_mo(&[
            (b"", HEADER_2PLURAL),
            (b"apple\0apples", b"pomme\0pommes"),
        ]);
        let cat = Catalog::parse(&mo).unwrap();
        assert_eq!(cat.ngettext("apple", "apples", 0), "pommes");
        assert_eq!(cat.ngettext("apple", "apples", 1), "pomme");
        assert_eq!(cat.ngettext("apple", "apples", 2), "pommes");
    }

    #[test]
    fn test_ngettext_clamps_index_to_available_forms() {
        // rule may yield an index beyond the forms actually present in the msgstr
        let mo = build_mo(&[
            (b"", b"Plural-Forms: nplurals=3; plural=n;\n"),
            (b"f", b"a\0b"),
        ]);
        let cat = Catalog::parse(&mo).unwrap();
        assert_eq!(cat.ngettext("f", "fs", 0), "a");
        assert_eq!(cat.ngettext("f", "fs", 1), "b");
        // plural_index(2) == 2 but only 2 forms exist -> clamped to last form
        assert_eq!(cat.ngettext("f", "fs", 2), "b");
    }

    #[test]
    fn test_ngettext_default_rule_without_header() {
        // catalog without Plural-Forms header: default rule is (n != 1)
        let mo = build_mo(&[(b"", b"Content-Type: text/plain\n"), (b"a\0as", b"x\0y")]);
        let cat = Catalog::parse(&mo).unwrap();
        assert_eq!(cat.ngettext("a", "as", 1), "x");
        assert_eq!(cat.ngettext("a", "as", 0), "y");
        assert_eq!(cat.ngettext("a", "as", 7), "y");
    }

    #[test]
    fn test_ngettext_nplurals_zero_does_not_panic() {
        // nplurals=0 would make `idx % nplurals` a division by zero without the guard
        let mo = build_mo(&[(b"", b"Plural-Forms: nplurals=0; plural=n;\n"), (b"a", b"x")]);
        let cat = Catalog::parse(&mo).unwrap();
        assert_eq!(cat.nplurals, 0);
        assert_eq!(cat.ngettext("a", "as", 5), "x");
    }

    #[test]
    fn test_ngettext_missing_msgid_falls_back_to_singular_or_plural() {
        let mo = build_mo(&[(b"", HEADER_2PLURAL), (b"apple", b"pomme")]);
        let cat = Catalog::parse(&mo).unwrap();
        assert_eq!(cat.ngettext("cherry", "cherries", 1), "cherry");
        assert_eq!(cat.ngettext("cherry", "cherries", 2), "cherries");
        assert_eq!(cat.ngettext("cherry", "cherries", 0), "cherries");
        // identity fallback also holds for an empty catalog
        let empty = Catalog::default();
        assert_eq!(empty.ngettext("a", "b", 1), "a");
        assert_eq!(empty.ngettext("a", "b", 9), "b");
    }

    #[test]
    fn test_npgettext_contextual_plural_forms() {
        let mo = build_mo(&[
            (b"", HEADER_2PLURAL),
            (b"fruit\x04apple\0apples", b"pomme\0pommes"),
        ]);
        let cat = Catalog::parse(&mo).unwrap();
        assert_eq!(cat.npgettext("fruit", "apple", "apples", 1), "pomme");
        assert_eq!(cat.npgettext("fruit", "apple", "apples", 3), "pommes");
    }

    #[test]
    fn test_npgettext_missing_falls_back_to_singular_or_plural() {
        let mo = build_mo(&[(b"", HEADER_2PLURAL), (b"fruit\x04apple", b"pomme")]);
        let cat = Catalog::parse(&mo).unwrap();
        assert_eq!(cat.npgettext("veg", "apple", "apples", 1), "apple");
        assert_eq!(cat.npgettext("veg", "apple", "apples", 4), "apples");
    }

    // ---------------- PluralExpr::parse ----------------

    #[test]
    fn test_plural_expr_parse_rejects_invalid_input() {
        assert!(PluralExpr::parse("").is_none());
        assert!(PluralExpr::parse("n ?").is_none()); // incomplete ternary
        assert!(PluralExpr::parse("n ? 1 :").is_none()); // missing else branch
        assert!(PluralExpr::parse("n ==").is_none()); // missing right operand
        assert!(PluralExpr::parse("(n").is_none()); // unbalanced paren
        assert!(PluralExpr::parse("1 2").is_none()); // trailing garbage
        assert!(PluralExpr::parse("1 < n < 3").is_none()); // comparisons don't chain
        assert!(PluralExpr::parse("-1").is_none()); // no negative literals
        assert!(PluralExpr::parse("n & 1").is_none()); // single & is not an operator
    }

    #[test]
    fn test_plural_expr_parse_tolerates_whitespace() {
        let e = PluralExpr::parse("  n\t!=\n1 ").unwrap();
        assert_eq!(e.eval(1), 0);
        assert_eq!(e.eval(2), 1);
    }

    #[test]
    fn test_plural_expr_num_constant_ignores_n() {
        let e = PluralExpr::parse("42").unwrap();
        assert_eq!(e.eval(0), 42);
        assert_eq!(e.eval(100), 42);
    }

    #[test]
    fn test_plural_expr_comparison_operators() {
        let cases: &[(&str, i64, i64)] = &[
            ("n == 1", 1, 1),
            ("n != 1", 1, 0),
            ("n < 2", 1, 1),
            ("n <= 1", 2, 0),
            ("n > 0", 1, 1),
            ("n >= 2", 1, 0),
        ];
        for (src, n, expected) in cases {
            assert_eq!(PluralExpr::parse(src).unwrap().eval(*n), *expected, "{src}");
        }
    }

    #[test]
    fn test_plural_expr_and_binds_tighter_than_or() {
        // `1 || 0 && 0` must parse as 1 || (0 && 0) == 1, not (1 || 0) && 0 == 0
        let e = PluralExpr::parse("1 || 0 && 0").unwrap();
        assert_eq!(e.eval(0), 1);
        let e = PluralExpr::parse("n < 2 || n > 10").unwrap();
        assert_eq!(e.eval(1), 1);
        assert_eq!(e.eval(5), 0);
        assert_eq!(e.eval(11), 1);
    }

    #[test]
    fn test_plural_expr_comparison_binds_tighter_than_and() {
        let e = PluralExpr::parse("n > 1 && n < 5").unwrap();
        assert_eq!(e.eval(1), 0);
        assert_eq!(e.eval(3), 1);
        assert_eq!(e.eval(5), 0);
    }

    #[test]
    fn test_plural_expr_mod_binds_tighter_than_comparison() {
        let e = PluralExpr::parse("n % 2 == 0").unwrap();
        assert_eq!(e.eval(4), 1);
        assert_eq!(e.eval(3), 0);
    }

    #[test]
    fn test_plural_expr_ternary_lowest_precedence_and_right_assoc() {
        // condition is a full || expression, branches chain right-associatively
        let e = PluralExpr::parse("n == 1 || n == 2 ? 10 : n == 3 ? 20 : 30").unwrap();
        assert_eq!(e.eval(1), 10);
        assert_eq!(e.eval(2), 10);
        assert_eq!(e.eval(3), 20);
        assert_eq!(e.eval(4), 30);
    }

    #[test]
    fn test_plural_expr_not_operator() {
        let e = PluralExpr::parse("!n").unwrap();
        assert_eq!(e.eval(0), 1);
        assert_eq!(e.eval(3), 0);
        let e = PluralExpr::parse("!!n").unwrap();
        assert_eq!(e.eval(0), 0);
        assert_eq!(e.eval(3), 1);
        let e = PluralExpr::parse("!(n == 1)").unwrap();
        assert_eq!(e.eval(1), 0);
        assert_eq!(e.eval(2), 1);
    }

    #[test]
    fn test_plural_expr_add_sub() {
        let e = PluralExpr::parse("n + 2 - 1").unwrap();
        assert_eq!(e.eval(5), 6);
    }

    #[test]
    fn test_plural_expr_mod_by_zero_returns_zero() {
        // guard against division by zero: x % 0 evaluates to 0 instead of panicking
        let e = PluralExpr::parse("n % (n - n)").unwrap();
        assert_eq!(e.eval(7), 0);
        let e = PluralExpr::parse("n % 0 ? 1 : 2").unwrap();
        assert_eq!(e.eval(7), 2);
    }

    #[test]
    fn test_plural_expr_mod_binds_tighter_than_add() {
        // C/gettext precedence: `%` is multiplicative, so `1 + n % 2`
        // parses as `1 + (n % 2)`, not `(1 + n) % 2`.
        let e = PluralExpr::parse("1 + n % 2").unwrap();
        assert_eq!(e.eval(1), 2);
        assert_eq!(e.eval(2), 1);
        // left-associative within the multiplicative level
        let e = PluralExpr::parse("n % 5 % 3").unwrap();
        assert_eq!(e.eval(7), 2); // (7 % 5) % 3
    }

    #[test]
    fn test_plural_expr_parentheses_override_precedence() {
        let e = PluralExpr::parse("1 + (n % 2)").unwrap();
        assert_eq!(e.eval(1), 2);
        assert_eq!(e.eval(2), 1);
    }
}
