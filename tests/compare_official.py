"""Compare official docxtpl test outputs between reference engine and docxtplrs.

Usage: python3 tests/compare_official.py /tmp/dtpl-ref/tests/output /tmp/dtpl-rs/tests/output
"""

import hashlib
import re
import sys
import zipfile


def features(path):
    try:
        z = zipfile.ZipFile(path)
    except Exception as e:
        return {"error": str(e)}
    out = {}
    for name in sorted(z.namelist()):
        data = z.read(name)
        if name.endswith(".xml") or name.endswith(".rels"):
            try:
                xml = data.decode("utf-8")
            except UnicodeDecodeError:
                continue
            text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))
            f = {
                "text": text,
                "n_p": len(re.findall(r"<w:p[ >]", xml)),
                "n_tr": len(re.findall(r"<w:tr[ >]", xml)),
                "n_tc": len(re.findall(r"<w:tc[ >]", xml)),
                "gridspan": sorted(re.findall(r'w:gridSpan w:val="(\d+)"', xml)),
                "vmerge": sorted(re.findall(r'w:vMerge w:val="(\w+)"', xml)),
                "shd": sorted(re.findall(r'<w:shd[^>]*w:fill="([^"]+)"', xml)),
                "br": xml.count("<w:br/>") + len(re.findall(r'<w:br [^/]*/>', xml)),
                "tab": xml.count("<w:tab/>"),
                "drawing": xml.count("<w:drawing>"),
                "hyperlink_rel": sorted(re.findall(r'Target="(https?://[^"]+)"', xml)),
                "img_rel": len(re.findall(r'relationships/image"', xml)),
                "gridcols": [int(w) for w in re.findall(r'<w:gridCol w:w="(\d+)"/>', xml)],
            }
            out[name] = {k: v for k, v in f.items() if v not in (0, [], "")}
        else:
            # binary part: compare content hash
            out[name] = {"sha1": hashlib.sha1(data).hexdigest()[:12]}
    return out


def main():
    import glob
    import os

    ref_dir, rs_dir = sys.argv[1], sys.argv[2]
    ref_files = {os.path.basename(p): p for p in glob.glob(ref_dir + "/*.docx")}
    rs_files = {os.path.basename(p): p for p in glob.glob(rs_dir + "/*.docx")}

    only_ref = sorted(set(ref_files) - set(rs_files))
    only_rs = sorted(set(rs_files) - set(ref_files))
    if only_ref:
        print("ONLY IN REF (skipped):", ", ".join(only_ref))
    if only_rs:
        print("ONLY IN RS:", ", ".join(only_rs))

    n_ok = n_bad = 0
    for name in sorted(set(ref_files) & set(rs_files)):
        a = features(ref_files[name])
        b = features(rs_files[name])
        # compare per-part
        diffs = []
        for part in sorted(set(a) | set(b)):
            fa, fb = a.get(part, {}), b.get(part, {})
            if fa != fb:
                keys = set(fa) | set(fb)
                for k in sorted(keys):
                    if fa.get(k) != fb.get(k):
                        diffs.append((part, k, fa.get(k), fb.get(k)))
        if diffs:
            n_bad += 1
            print(f"\n=== {name} ===")
            for part, k, va, vb in diffs[:12]:
                sa, sb = repr(va), repr(vb)
                if len(sa) > 120:
                    sa = sa[:120] + "..."
                if len(sb) > 120:
                    sb = sb[:120] + "..."
                print(f"  [{part}] {k}:\n    ref: {sa}\n    rs:  {sb}")
        else:
            n_ok += 1
    print(f"\n{n_ok} files match, {n_bad} files differ")
    return 1 if n_bad else 0


if __name__ == "__main__":
    sys.exit(main())
