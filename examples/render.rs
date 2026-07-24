//! Rust API demo: render a docx template without Python.
//!
//! Run: cargo run --example render --release -- template.docx out.docx

use docxtplrs::template::TplCore;
use minijinja::Value;
use std::collections::BTreeMap;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let template = args.next().unwrap_or_else(|| "template.docx".into());
    let output = args.next().unwrap_or_else(|| "out.docx".into());

    // build the rendering context as a minijinja value
    let mut ctx = BTreeMap::new();
    ctx.insert("title".to_string(), Value::from("Quarterly Report"));
    ctx.insert(
        "items".to_string(),
        Value::from_serialize(&vec!["alpha", "beta", "gamma"]),
    );

    // load the template and render (autoescape = false)
    let mut tpl = TplCore::new(std::fs::read(&template)?);
    tpl.render(false, &|_core, _part| Ok(Value::from_serialize(&ctx)))?;

    // write the result
    let bytes = tpl.save_bytes()?;
    std::fs::write(&output, &bytes)?;
    println!("saved {} ({} bytes)", output, bytes.len());

    // template introspection
    println!("variables: {:?}", tpl.undeclared_variables(None)?);
    Ok(())
}
