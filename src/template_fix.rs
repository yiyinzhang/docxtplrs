    /// Process {%p ... %} tags (paragraph-level)
    fn process_paragraph_tags(&self, content: &str) -> Result<String> {
        let mut result = content.to_string();

        // Simply replace {%p ... %} with {% ... %}
        let re = Regex::new(r"\{%p\s+(.+?)%\}")?;
        result = re.replace_all(&result, |caps: &regex::Captures| {
            format!("{{% {} %}}", &caps[1])
        }).to_string();

        Ok(result)
    }

    /// Process {%tr ... %} tags (table row level)
    fn process_table_row_tags(&self, content: &str) -> Result<String> {
        let mut result = content.to_string();

        // Simply replace {%tr ... %} with {% ... %}
        let re = Regex::new(r"\{%tr\s+(.+?)%\}")?;
        result = re.replace_all(&result, |caps: &regex::Captures| {
            format!("{{% {} %}}", &caps[1])
        }).to_string();

        Ok(result)
    }

    /// Process {%tc ... %} tags (table cell level)
    fn process_table_cell_tags(&self, content: &str) -> Result<String> {
        let mut result = content.to_string();

        // Simply replace {%tc ... %} with {% ... %}
        let re = Regex::new(r"\{%tc\s+(.+?)%\}")?;
        result = re.replace_all(&result, |caps: &regex::Captures| {
            format!("{{% {} %}}", &caps[1])
        }).to_string();

        Ok(result)
    }

    /// Process {%r ... %} tags (run level)
    fn process_run_tags(&self, content: &str) -> Result<String> {
        let mut result = content.to_string();

        // Replace {%r ... %} with {% ... %}
        let re = Regex::new(r"\{%r\s+(.+?)%\}")?;
        result = re.replace_all(&result, |caps: &regex::Captures| {
            format!("{{% {} %}}", &caps[1])
        }).to_string();

        // Also handle {{r ... }} for variables - replace with {{ ... }}
        let re2 = Regex::new(r"\{\{r\s+(.+?)\}\}")?;
        result = re2.replace_all(&result, |caps: &regex::Captures| {
            format!("{{{{ {} }}}}", &caps[1])
        }).to_string();

        Ok(result)
    }
