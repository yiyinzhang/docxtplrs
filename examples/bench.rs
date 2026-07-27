//! Stage-level benchmark for the render pipeline.
//!
//! Run: cargo run --example bench --release

use docxtplrs::template::{fix_tables_and_docpr, render_xml_str, TplCore};
use minijinja::Value;
use std::alloc::{GlobalAlloc, Layout, System};
use std::collections::BTreeMap;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::time::Instant;

struct Counting;
static ALLOCS: AtomicUsize = AtomicUsize::new(0);
static BYTES: AtomicUsize = AtomicUsize::new(0);
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, l: Layout) -> *mut u8 {
        ALLOCS.fetch_add(1, Ordering::Relaxed);
        BYTES.fetch_add(l.size(), Ordering::Relaxed);
        System.alloc(l)
    }
    unsafe fn dealloc(&self, p: *mut u8, l: Layout) {
        System.dealloc(p, l)
    }
}
#[global_allocator]
static A: Counting = Counting;

fn alloc_stats() -> (usize, usize) {
    (
        ALLOCS.swap(0, Ordering::Relaxed),
        BYTES.swap(0, Ordering::Relaxed),
    )
}
fn report(stage: &str, t0: Instant, a0: (usize, usize)) {
    let (a1, b1) = alloc_stats();
    println!(
        "{stage:20} {:>9.1?}  allocs={:>8} ({} new)  alloc_bytes={:.1}MB",
        t0.elapsed(),
        a1,
        a1.saturating_sub(a0.0),
        b1 as f64 / 1e6
    );
}

fn make_ctx() -> Value {
    let rows: Vec<BTreeMap<String, Value>> = (0..10)
        .map(|i| {
            let mut m = BTreeMap::new();
            m.insert("name".to_string(), Value::from(format!("rowname-{i}")));
            m.insert("value".to_string(), Value::from(i * 7));
            m
        })
        .collect();
    let groups: Vec<BTreeMap<String, Value>> = (0..3)
        .map(|g| {
            let mut m = BTreeMap::new();
            m.insert("title".to_string(), Value::from(format!("group-{g}")));
            m.insert(
                "entries".to_string(),
                Value::from_serialize(&(0..10).map(|i| format!("item-{i}")).collect::<Vec<_>>()),
            );
            m
        })
        .collect();
    let mut user = BTreeMap::new();
    user.insert("name".to_string(), Value::from("张三 Zhang"));

    let mut ctx = BTreeMap::new();
    ctx.insert("user".to_string(), Value::from_serialize(&user));
    ctx.insert("rows".to_string(), Value::from_serialize(&rows));
    ctx.insert("groups".to_string(), Value::from_serialize(&groups));
    Value::from_serialize(&ctx)
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = std::fs::read("/tmp/bench/big.docx")?;
    let ctx = make_ctx();

    // ---- end-to-end ----
    let t = Instant::now();
    let mut tpl = TplCore::new(bytes.clone());
    let mut ctx2 = ctx.clone();
    tpl.render(false, &|_c, _p| Ok(ctx2.clone()))?;
    println!("e2e render:          {:?}", t.elapsed());
    let t = Instant::now();
    let out = tpl.save_bytes()?;
    println!("save_bytes:          {:?} ({} bytes)", t.elapsed(), out.len());

    // ---- stage-level on raw document.xml ----
    let z = std::fs::File::open("/tmp/bench/big.docx")?;
    let mut zip = zip::ZipArchive::new(z)?;
    let xml = {
        use std::io::Read;
        let mut s = String::new();
        zip.by_name("word/document.xml")?.read_to_string(&mut s)?;
        s
    };
    println!("document.xml:        {} bytes", xml.len());

    let t = Instant::now();
    let a = alloc_stats();
    let patched = docxtplrs::patch::patch_xml(&xml);
    report("patch_xml", t, a);
    println!("  -> {} bytes", patched.len());

    let mut core = TplCore::new(bytes);
    let t = Instant::now();
    let a = alloc_stats();
    let rendered = render_xml_str(&patched, ctx.clone(), false, &mut core)?;
    report("minijinja render", t, a);
    println!("  -> {} bytes", rendered.len());

    let t = Instant::now();
    let a = alloc_stats();
    let mut idx = 1000u32;
    let fixed = fix_tables_and_docpr(&rendered, &mut idx)?;
    report("fix_tables", t, a);
    println!("  -> {} bytes", fixed.len());

    let t = Instant::now();
    let a = alloc_stats();
    let dom = docxtplrs::xmldom::Document::parse(&rendered)?;
    report("xmldom parse", t, a);
    let t = Instant::now();
    let a = alloc_stats();
    let ser = dom.serialize();
    report("xmldom serialize", t, a);
    println!("  -> {} bytes", ser.len());

    ctx2 = ctx.clone();
    let _ = &mut ctx2;
    Ok(())
}
