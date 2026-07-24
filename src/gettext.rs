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
        let mut left = self.parse_unary()?;
        loop {
            self.skip_ws();
            match self.peek() {
                Some('+') => {
                    self.next();
                    let right = self.parse_unary()?;
                    left = PluralExpr::Add(Box::new(left), Box::new(right));
                }
                Some('-') => {
                    self.next();
                    let right = self.parse_unary()?;
                    left = PluralExpr::Sub(Box::new(left), Box::new(right));
                }
                Some('%') => {
                    self.next();
                    let right = self.parse_unary()?;
                    left = PluralExpr::Mod(Box::new(left), Box::new(right));
                }
                _ => return Some(left),
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
}
