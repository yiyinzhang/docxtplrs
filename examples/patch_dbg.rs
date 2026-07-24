fn main() {
    let xml = std::fs::read_to_string("/tmp/treport_doc.xml").unwrap();
    let t = std::time::Instant::now();
    let out = std::panic::catch_unwind(|| docxtplrs::patch::patch_xml(&xml));
    match out {
        Ok(s) => println!("patch_xml ok in {:?}, {} bytes", t.elapsed(), s.len()),
        Err(_) => println!("patch_xml PANIC in {:?}", t.elapsed()),
    }

    if std::env::var("SMALL").is_ok() {
        let small = r#"<w:body><w:p><w:r><w:t>{%p for x in items %}</w:t></w:r></w:p><w:p><w:r><w:t>item {{ x }}</w:t></w:r></w:p><w:p><w:r><w:t>{%p endfor %}</w:t></w:r></w:p></w:body>"#;
        println!("--- small p-loop ---");
        println!("{}", docxtplrs::patch::patch_xml(small));
        let small_r = r#"<w:body><w:p><w:r><w:t>{{r rt }}</w:t></w:r></w:p></w:body>"#;
        println!("--- small r-run ---");
        println!("{}", docxtplrs::patch::patch_xml(small_r));
    }
}
