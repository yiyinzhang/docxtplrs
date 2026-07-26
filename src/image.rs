//! Image header parsing: pixel dimensions, DPI, content type, extension.
//! Mirrors python-docx image handling (default 72 DPI when absent).

#[derive(Debug, Clone)]
pub struct ImageInfo {
    pub px_width: u32,
    pub px_height: u32,
    pub horz_dpi: u32,
    pub vert_dpi: u32,
    pub content_type: &'static str,
    pub default_ext: &'static str,
}

pub const EMU_PER_INCH: i64 = 914400;

impl ImageInfo {
    pub fn parse(blob: &[u8]) -> Result<ImageInfo, String> {
        if blob.len() >= 8 && &blob[0..8] == b"\x89PNG\r\n\x1a\n" {
            return parse_png(blob);
        }
        if blob.len() >= 2 && blob[0] == 0xFF && blob[1] == 0xD8 {
            return parse_jpeg(blob);
        }
        if blob.len() >= 6 && (&blob[0..6] == b"GIF87a" || &blob[0..6] == b"GIF89a") {
            return parse_gif(blob);
        }
        if blob.len() >= 2 && &blob[0..2] == b"BM" {
            return parse_bmp(blob);
        }
        if blob.len() >= 4 && (&blob[0..4] == b"II*\x00" || &blob[0..4] == b"MM\x00*") {
            return parse_tiff(blob);
        }
        Err("UnrecognizedImageError: cannot identify image format".to_string())
    }

    /// Native width in EMU
    pub fn width_emu(&self) -> i64 {
        (EMU_PER_INCH as f64 * self.px_width as f64 / self.horz_dpi as f64).round() as i64
    }

    pub fn height_emu(&self) -> i64 {
        (EMU_PER_INCH as f64 * self.px_height as f64 / self.vert_dpi as f64).round() as i64
    }

    /// python-docx Image.scaled_dimensions
    pub fn scaled_dimensions(&self, width: Option<i64>, height: Option<i64>) -> (i64, i64) {
        let native_w = self.width_emu();
        let native_h = self.height_emu();
        match (width, height) {
            (Some(w), Some(h)) => (w, h),
            (None, Some(h)) => {
                let factor = h as f64 / native_h as f64;
                ((native_w as f64 * factor).round() as i64, h)
            }
            (Some(w), None) => {
                let factor = w as f64 / native_w as f64;
                (w, (native_h as f64 * factor).round() as i64)
            }
            (None, None) => (native_w, native_h),
        }
    }
}

fn le_u32(b: &[u8], off: usize) -> u32 {
    u32::from_le_bytes([b[off], b[off + 1], b[off + 2], b[off + 3]])
}
fn be_u32(b: &[u8], off: usize) -> u32 {
    u32::from_be_bytes([b[off], b[off + 1], b[off + 2], b[off + 3]])
}
fn be_u16(b: &[u8], off: usize) -> u16 {
    u16::from_be_bytes([b[off], b[off + 1]])
}

fn parse_png(blob: &[u8]) -> Result<ImageInfo, String> {
    // IHDR starts at offset 8: len(4) "IHDR" width(4) height(4)
    if blob.len() < 24 || &blob[12..16] != b"IHDR" {
        return Err("invalid PNG".into());
    }
    let px_width = be_u32(blob, 16);
    let px_height = be_u32(blob, 20);
    let mut horz_dpi = 72u32;
    let mut vert_dpi = 72u32;
    // scan chunks for pHYs
    let mut pos = 8usize;
    while pos + 12 <= blob.len() {
        let len = be_u32(blob, pos) as usize;
        let ctype = &blob[pos + 4..pos + 8];
        if ctype == b"pHYs" && pos + 12 + 9 <= blob.len() {
            let ppux = be_u32(blob, pos + 8);
            let ppuy = be_u32(blob, pos + 12);
            let unit = blob[pos + 16];
            if unit == 1 && ppux > 0 && ppuy > 0 {
                // pixels per meter -> dpi
                horz_dpi = (ppux as f64 * 0.0254).round() as u32;
                vert_dpi = (ppuy as f64 * 0.0254).round() as u32;
            }
            break;
        }
        pos += 12 + len;
        if ctype == b"IDAT" {
            break;
        }
    }
    Ok(ImageInfo {
        px_width,
        px_height,
        horz_dpi,
        vert_dpi,
        content_type: "image/png",
        default_ext: "png",
    })
}

fn parse_jpeg(blob: &[u8]) -> Result<ImageInfo, String> {
    let mut pos = 2usize;
    let mut px_width = 0u32;
    let mut px_height = 0u32;
    let mut horz_dpi = 72u32;
    let mut vert_dpi = 72u32;
    while pos + 4 <= blob.len() {
        if blob[pos] != 0xFF {
            pos += 1;
            continue;
        }
        let marker = blob[pos + 1];
        // markers without length
        if marker == 0xD8 || marker == 0xD9 || (0xD0..=0xD7).contains(&marker) || marker == 0x01 {
            pos += 2;
            continue;
        }
        if pos + 4 > blob.len() {
            break;
        }
        let seg_len = be_u16(blob, pos + 2) as usize;
        if seg_len < 2 {
            break;
        }
        match marker {
            0xE0 => {
                // APP0 JFIF
                if pos + 16 <= blob.len() && &blob[pos + 4..pos + 9] == b"JFIF\x00" {
                    let units = blob[pos + 11];
                    let xden = be_u16(blob, pos + 12) as u32;
                    let yden = be_u16(blob, pos + 14) as u32;
                    if xden > 0 && yden > 0 {
                        match units {
                            1 => {
                                horz_dpi = xden;
                                vert_dpi = yden;
                            }
                            2 => {
                                horz_dpi = (xden as f64 * 2.54).round() as u32;
                                vert_dpi = (yden as f64 * 2.54).round() as u32;
                            }
                            _ => {}
                        }
                    }
                }
            }
            0xC0..=0xC3 | 0xC5..=0xC7 | 0xC9..=0xCB | 0xCD..=0xCF => {
                // SOFn
                if pos + 9 <= blob.len() {
                    px_height = be_u16(blob, pos + 5) as u32;
                    px_width = be_u16(blob, pos + 7) as u32;
                }
                break;
            }
            _ => {}
        }
        pos += 2 + seg_len;
    }
    if px_width == 0 || px_height == 0 {
        return Err("invalid JPEG: no SOF found".into());
    }
    Ok(ImageInfo {
        px_width,
        px_height,
        horz_dpi,
        vert_dpi,
        content_type: "image/jpeg",
        default_ext: "jpg",
    })
}

fn parse_gif(blob: &[u8]) -> Result<ImageInfo, String> {
    if blob.len() < 10 {
        return Err("invalid GIF".into());
    }
    let px_width = u16::from_le_bytes([blob[6], blob[7]]) as u32;
    let px_height = u16::from_le_bytes([blob[8], blob[9]]) as u32;
    Ok(ImageInfo {
        px_width,
        px_height,
        horz_dpi: 72,
        vert_dpi: 72,
        content_type: "image/gif",
        default_ext: "gif",
    })
}

fn parse_bmp(blob: &[u8]) -> Result<ImageInfo, String> {
    if blob.len() < 26 {
        return Err("invalid BMP".into());
    }
    let dib_size = le_u32(blob, 14) as usize;
    let (px_width, px_height);
    let mut horz_dpi = 72u32;
    let mut vert_dpi = 72u32;
    if dib_size == 12 {
        // BITMAPCOREHEADER
        px_width = u16::from_le_bytes([blob[18], blob[19]]) as u32;
        px_height = u16::from_le_bytes([blob[20], blob[21]]) as u32;
    } else {
        px_width = le_u32(blob, 18);
        // height is signed: negative means top-down row order
        px_height = (le_u32(blob, 22) as i32).unsigned_abs();
        // BITMAPINFOHEADER: biXPelsPerMeter @38, biYPelsPerMeter @42
        if blob.len() >= 46 {
            let xppm = le_u32(blob, 38);
            let yppm = le_u32(blob, 42);
            if xppm > 0 && yppm > 0 {
                horz_dpi = (xppm as f64 * 0.0254).round() as u32;
                vert_dpi = (yppm as f64 * 0.0254).round() as u32;
            }
        }
    }
    Ok(ImageInfo {
        px_width,
        px_height,
        horz_dpi,
        vert_dpi,
        content_type: "image/bmp",
        default_ext: "bmp",
    })
}

fn parse_tiff(blob: &[u8]) -> Result<ImageInfo, String> {
    let little = &blob[0..2] == b"II";
    let rd_u16 = |off: usize| -> u16 {
        if little {
            u16::from_le_bytes([blob[off], blob[off + 1]])
        } else {
            u16::from_be_bytes([blob[off], blob[off + 1]])
        }
    };
    let rd_u32 = |off: usize| -> u32 {
        if little {
            u32::from_le_bytes([blob[off], blob[off + 1], blob[off + 2], blob[off + 3]])
        } else {
            u32::from_be_bytes([blob[off], blob[off + 1], blob[off + 2], blob[off + 3]])
        }
    };
    if blob.len() < 8 {
        return Err("invalid TIFF".into());
    }
    let ifd_off = rd_u32(4) as usize;
    if ifd_off + 2 > blob.len() {
        return Err("invalid TIFF".into());
    }
    let count = rd_u16(ifd_off) as usize;
    let mut width = 0u32;
    let mut height = 0u32;
    let mut xres: Option<(u32, u32)> = None;
    let mut yres: Option<(u32, u32)> = None;
    let mut res_unit: u16 = 2;
    for i in 0..count {
        let e = ifd_off + 2 + i * 12;
        if e + 12 > blob.len() {
            break;
        }
        let tag = rd_u16(e);
        let typ = rd_u16(e + 2);
        let val = rd_u32(e + 8);
        // Tags 256/257 may be stored as SHORT (type 3, value inline in the
        // first 2 bytes of the value field) or LONG (type 4). Reading a SHORT
        // as u32 breaks big-endian files, where the value sits in the high
        // half of the field.
        let dim = if typ == 3 { rd_u16(e + 8) as u32 } else { val };
        match tag {
            256 => width = dim,
            257 => height = dim,
            282 => {
                let off = val as usize;
                if off + 8 <= blob.len() {
                    xres = Some((rd_u32(off), rd_u32(off + 4)));
                }
            }
            283 => {
                let off = val as usize;
                if off + 8 <= blob.len() {
                    yres = Some((rd_u32(off), rd_u32(off + 4)));
                }
            }
            296 => res_unit = rd_u16(e + 8),
            _ => {}
        }
    }
    if width == 0 || height == 0 {
        return Err("invalid TIFF: missing dimensions".into());
    }
    let to_dpi = |res: Option<(u32, u32)>| -> u32 {
        if let Some((num, den)) = res {
            if den > 0 && num > 0 {
                let r = num as f64 / den as f64;
                return match res_unit {
                    2 => r.round() as u32,
                    3 => (r * 2.54).round() as u32,
                    _ => 72,
                };
            }
        }
        72
    };
    Ok(ImageInfo {
        px_width: width,
        px_height: height,
        horz_dpi: to_dpi(xres),
        vert_dpi: to_dpi(yres),
        content_type: "image/tiff",
        default_ext: "tiff",
    })
}

/// python-docx Length helpers (EMU)
pub mod length {
    pub const EMU_PER_INCH: i64 = 914400;
    pub const EMU_PER_CM: i64 = 360000;
    pub const EMU_PER_MM: i64 = 36000;
    pub const EMU_PER_PT: i64 = 12700;
    pub const EMU_PER_TWIP: i64 = 635;

    pub fn inches(v: f64) -> i64 {
        (v * EMU_PER_INCH as f64) as i64
    }
    pub fn cm(v: f64) -> i64 {
        (v * EMU_PER_CM as f64) as i64
    }
    pub fn mm(v: f64) -> i64 {
        (v * EMU_PER_MM as f64) as i64
    }
    pub fn pt(v: f64) -> i64 {
        (v * EMU_PER_PT as f64) as i64
    }
    pub fn twips(v: f64) -> i64 {
        (v * EMU_PER_TWIP as f64) as i64
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // ---------- minimal byte-stream builders ----------

    /// Minimal PNG: 8-byte signature + IHDR chunk (CRC not validated by parser)
    /// plus an optional pHYs chunk.
    fn png_bytes(w: u32, h: u32, phys: Option<(u32, u32, u8)>) -> Vec<u8> {
        let mut v = Vec::new();
        v.extend_from_slice(b"\x89PNG\r\n\x1a\n");
        v.extend_from_slice(&13u32.to_be_bytes()); // IHDR data length
        v.extend_from_slice(b"IHDR");
        v.extend_from_slice(&w.to_be_bytes());
        v.extend_from_slice(&h.to_be_bytes());
        v.extend_from_slice(&[8, 6, 0, 0, 0]); // bit depth, color type, comp, filter, interlace
        v.extend_from_slice(&[0, 0, 0, 0]); // CRC (ignored)
        if let Some((ppux, ppuy, unit)) = phys {
            v.extend_from_slice(&9u32.to_be_bytes());
            v.extend_from_slice(b"pHYs");
            v.extend_from_slice(&ppux.to_be_bytes());
            v.extend_from_slice(&ppuy.to_be_bytes());
            v.push(unit);
            v.extend_from_slice(&[0, 0, 0, 0]); // CRC (ignored)
        }
        v
    }

    /// Minimal JPEG: SOI + optional APP0/JFIF segment + optional SOF0 segment.
    /// app0 = (units, xdensity, ydensity); sof = (height, width).
    fn jpeg_bytes(app0: Option<(u8, u16, u16)>, sof: Option<(u16, u16)>) -> Vec<u8> {
        let mut v = vec![0xFF, 0xD8];
        if let Some((units, xden, yden)) = app0 {
            v.extend_from_slice(&[0xFF, 0xE0]);
            v.extend_from_slice(&16u16.to_be_bytes()); // segment length incl. length field
            v.extend_from_slice(b"JFIF\x00");
            v.extend_from_slice(&[1, 1]); // JFIF version
            v.push(units);
            v.extend_from_slice(&xden.to_be_bytes());
            v.extend_from_slice(&yden.to_be_bytes());
            v.extend_from_slice(&[0, 0]); // thumbnail w/h
        }
        if let Some((h, w)) = sof {
            v.extend_from_slice(&[0xFF, 0xC0]);
            v.extend_from_slice(&8u16.to_be_bytes()); // length (no components needed here)
            v.push(8); // sample precision
            v.extend_from_slice(&h.to_be_bytes());
            v.extend_from_slice(&w.to_be_bytes());
        }
        v
    }

    fn gif_bytes(magic: &[u8; 6], w: u16, h: u16) -> Vec<u8> {
        let mut v = Vec::new();
        v.extend_from_slice(magic);
        v.extend_from_slice(&w.to_le_bytes());
        v.extend_from_slice(&h.to_le_bytes());
        v.extend_from_slice(&[0, 0, 0]); // packed fields, bg color, aspect ratio
        v
    }

    /// BMP with BITMAPINFOHEADER (dib size 40), signed height, ppm resolution fields.
    fn bmp_info_bytes(w: i32, h: i32, xppm: u32, yppm: u32) -> Vec<u8> {
        let mut v = Vec::new();
        v.extend_from_slice(b"BM");
        v.extend_from_slice(&[0; 12]); // file size, reserved, pixel data offset
        v.extend_from_slice(&40u32.to_le_bytes()); // DIB header size
        v.extend_from_slice(&(w as u32).to_le_bytes()); // offset 18
        v.extend_from_slice(&h.to_le_bytes()); // offset 22 (signed)
        v.extend_from_slice(&[0; 12]); // planes, bpp, compression, image size
        v.extend_from_slice(&xppm.to_le_bytes()); // offset 38
        v.extend_from_slice(&yppm.to_le_bytes()); // offset 42
        v
    }

    /// BMP with BITMAPCOREHEADER (dib size 12), u16 dimensions.
    fn bmp_core_bytes(w: u16, h: u16) -> Vec<u8> {
        let mut v = Vec::new();
        v.extend_from_slice(b"BM");
        v.extend_from_slice(&[0; 12]);
        v.extend_from_slice(&12u32.to_le_bytes());
        v.extend_from_slice(&w.to_le_bytes());
        v.extend_from_slice(&h.to_le_bytes());
        v.extend_from_slice(&[0; 4]); // planes, bpp
        v
    }

    fn push16(v: &mut Vec<u8>, little: bool, x: u16) {
        if little {
            v.extend_from_slice(&x.to_le_bytes());
        } else {
            v.extend_from_slice(&x.to_be_bytes());
        }
    }
    fn push32(v: &mut Vec<u8>, little: bool, x: u32) {
        if little {
            v.extend_from_slice(&x.to_le_bytes());
        } else {
            v.extend_from_slice(&x.to_be_bytes());
        }
    }

    /// Minimal TIFF with one IFD at offset 8. Dimensions stored as LONG.
    /// res = Some((xres_num, yres_num, res_unit)) adds XResolution/YResolution
    /// (RATIONAL with denominator 1) and ResolutionUnit entries.
    fn tiff_bytes(little: bool, w: u32, h: u32, res: Option<(u32, u32, u16)>) -> Vec<u8> {
        let mut v = Vec::new();
        v.extend_from_slice(if little { b"II" } else { b"MM" });
        push16(&mut v, little, 42);
        push32(&mut v, little, 8); // IFD offset
        let n: u16 = if res.is_some() { 5 } else { 2 };
        push16(&mut v, little, n);
        // rational data lives right after IFD: 8 + 2 + n*12 + 4
        let data_off = 8 + 2 + n as u32 * 12 + 4;
        // tag 256 ImageWidth (LONG)
        push16(&mut v, little, 256);
        push16(&mut v, little, 4);
        push32(&mut v, little, 1);
        push32(&mut v, little, w);
        // tag 257 ImageLength (LONG)
        push16(&mut v, little, 257);
        push16(&mut v, little, 4);
        push32(&mut v, little, 1);
        push32(&mut v, little, h);
        if let Some((_, _, unit)) = res {
            // tag 282 XResolution (RATIONAL -> offset)
            push16(&mut v, little, 282);
            push16(&mut v, little, 5);
            push32(&mut v, little, 1);
            push32(&mut v, little, data_off);
            // tag 283 YResolution (RATIONAL -> offset)
            push16(&mut v, little, 283);
            push16(&mut v, little, 5);
            push32(&mut v, little, 1);
            push32(&mut v, little, data_off + 8);
            // tag 296 ResolutionUnit (SHORT, inline)
            push16(&mut v, little, 296);
            push16(&mut v, little, 3);
            push32(&mut v, little, 1);
            push16(&mut v, little, unit);
            push16(&mut v, little, 0);
        }
        push32(&mut v, little, 0); // next IFD offset
        if let Some((xres, yres, _)) = res {
            push32(&mut v, little, xres);
            push32(&mut v, little, 1);
            push32(&mut v, little, yres);
            push32(&mut v, little, 1);
        }
        v
    }

    fn info(px_w: u32, px_h: u32, hdpi: u32, vdpi: u32) -> ImageInfo {
        ImageInfo {
            px_width: px_w,
            px_height: px_h,
            horz_dpi: hdpi,
            vert_dpi: vdpi,
            content_type: "image/x-test",
            default_ext: "t",
        }
    }

    // ---------- format dispatch ----------

    #[test]
    fn test_parse_unrecognized_magic_errors() {
        let blob = b"\x00\x01\x02\x03\x04\x05\x06\x07\x08\x09";
        let err = ImageInfo::parse(blob).unwrap_err();
        assert!(err.contains("UnrecognizedImageError"), "got: {err}");
    }

    #[test]
    fn test_parse_empty_blob_errors() {
        assert!(ImageInfo::parse(&[]).is_err());
    }

    // ---------- PNG ----------

    #[test]
    fn test_parse_png_dimensions_and_defaults() {
        let blob = png_bytes(640, 480, None);
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.px_width, 640);
        assert_eq!(info.px_height, 480);
        assert_eq!(info.horz_dpi, 72); // no pHYs -> default 72
        assert_eq!(info.vert_dpi, 72);
        assert_eq!(info.content_type, "image/png");
        assert_eq!(info.default_ext, "png");
    }

    #[test]
    fn test_parse_png_phys_meter_unit_sets_dpi() {
        // 5906 px/m * 0.0254 = 150.01 -> 150 dpi
        let blob = png_bytes(10, 20, Some((5906, 5906, 1)));
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.horz_dpi, 150);
        assert_eq!(info.vert_dpi, 150);
    }

    #[test]
    fn test_parse_png_phys_non_meter_unit_keeps_default() {
        // unit != 1 (unknown unit) must not change dpi
        let blob = png_bytes(10, 20, Some((5906, 5906, 0)));
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.horz_dpi, 72);
        assert_eq!(info.vert_dpi, 72);
    }

    #[test]
    fn test_parse_png_phys_zero_density_keeps_default() {
        let blob = png_bytes(10, 20, Some((0, 5906, 1)));
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.horz_dpi, 72);
        assert_eq!(info.vert_dpi, 72);
    }

    #[test]
    fn test_parse_png_truncated_before_ihdr_errors() {
        let blob = b"\x89PNG\r\n\x1a\n\x00\x00"; // signature + 2 bytes only
        let err = ImageInfo::parse(blob).unwrap_err();
        assert_eq!(err, "invalid PNG");
    }

    #[test]
    fn test_parse_png_missing_ihdr_signature_errors() {
        // 24+ bytes but chunk type at offset 12 is not "IHDR"
        let mut blob = png_bytes(1, 1, None);
        blob[12..16].copy_from_slice(b"XXXX");
        assert_eq!(ImageInfo::parse(&blob).unwrap_err(), "invalid PNG");
    }

    // ---------- JPEG ----------

    #[test]
    fn test_parse_jpeg_dimensions_default_dpi() {
        let blob = jpeg_bytes(None, Some((480, 640))); // SOF: height, width
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.px_width, 640);
        assert_eq!(info.px_height, 480);
        assert_eq!(info.horz_dpi, 72);
        assert_eq!(info.vert_dpi, 72);
        assert_eq!(info.content_type, "image/jpeg");
        assert_eq!(info.default_ext, "jpg");
    }

    #[test]
    fn test_parse_jpeg_app0_dpi_units() {
        let blob = jpeg_bytes(Some((1, 150, 300)), Some((10, 20)));
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.horz_dpi, 150);
        assert_eq!(info.vert_dpi, 300);
    }

    #[test]
    fn test_parse_jpeg_app0_dpcm_units() {
        // units=2 -> dots per cm; 100 dpcm * 2.54 = 254 dpi
        let blob = jpeg_bytes(Some((2, 100, 50)), Some((10, 20)));
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.horz_dpi, 254);
        assert_eq!(info.vert_dpi, 127);
    }

    #[test]
    fn test_parse_jpeg_app0_zero_density_keeps_default() {
        let blob = jpeg_bytes(Some((1, 0, 0)), Some((10, 20)));
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.horz_dpi, 72);
        assert_eq!(info.vert_dpi, 72);
    }

    #[test]
    fn test_parse_jpeg_no_sof_errors() {
        let blob = jpeg_bytes(Some((1, 96, 96)), None);
        assert_eq!(
            ImageInfo::parse(&blob).unwrap_err(),
            "invalid JPEG: no SOF found"
        );
    }

    #[test]
    fn test_parse_jpeg_truncated_to_soi_errors() {
        let blob = [0xFF, 0xD8];
        assert!(ImageInfo::parse(&blob).is_err());
    }

    // ---------- GIF ----------

    #[test]
    fn test_parse_gif89a_dimensions() {
        let blob = gif_bytes(b"GIF89a", 320, 200);
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.px_width, 320);
        assert_eq!(info.px_height, 200);
        assert_eq!(info.horz_dpi, 72); // GIF has no DPI metadata
        assert_eq!(info.vert_dpi, 72);
        assert_eq!(info.content_type, "image/gif");
        assert_eq!(info.default_ext, "gif");
    }

    #[test]
    fn test_parse_gif87a_dimensions() {
        let blob = gif_bytes(b"GIF87a", 1, 65535);
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.px_width, 1);
        assert_eq!(info.px_height, 65535);
    }

    #[test]
    fn test_parse_gif_truncated_header_errors() {
        let blob = b"GIF89a\x01\x00"; // magic + partial width, < 10 bytes
        assert_eq!(ImageInfo::parse(blob).unwrap_err(), "invalid GIF");
    }

    // ---------- BMP ----------

    #[test]
    fn test_parse_bmp_info_header_with_ppm_dpi() {
        // 5906 px/m * 0.0254 = 150.01 -> 150 dpi
        let blob = bmp_info_bytes(100, 50, 5906, 5906);
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.px_width, 100);
        assert_eq!(info.px_height, 50);
        assert_eq!(info.horz_dpi, 150);
        assert_eq!(info.vert_dpi, 150);
        assert_eq!(info.content_type, "image/bmp");
        assert_eq!(info.default_ext, "bmp");
    }

    #[test]
    fn test_parse_bmp_zero_ppm_defaults_72() {
        let blob = bmp_info_bytes(8, 8, 0, 0);
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.horz_dpi, 72);
        assert_eq!(info.vert_dpi, 72);
    }

    #[test]
    fn test_parse_bmp_negative_height_is_absolute() {
        // negative height = top-down row order; parser must take abs value
        let blob = bmp_info_bytes(16, -24, 0, 0);
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.px_width, 16);
        assert_eq!(info.px_height, 24);
    }

    #[test]
    fn test_parse_bmp_core_header_u16_dimensions() {
        let blob = bmp_core_bytes(12, 34);
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.px_width, 12);
        assert_eq!(info.px_height, 34);
        assert_eq!(info.horz_dpi, 72); // core header has no resolution fields
    }

    #[test]
    fn test_parse_bmp_truncated_errors() {
        let blob = b"BM\x00\x00\x00";
        assert_eq!(ImageInfo::parse(blob).unwrap_err(), "invalid BMP");
    }

    // ---------- TIFF ----------

    #[test]
    fn test_parse_tiff_little_endian_with_resolution() {
        let blob = tiff_bytes(true, 800, 600, Some((300, 200, 2)));
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.px_width, 800);
        assert_eq!(info.px_height, 600);
        assert_eq!(info.horz_dpi, 300);
        assert_eq!(info.vert_dpi, 200);
        assert_eq!(info.content_type, "image/tiff");
        assert_eq!(info.default_ext, "tiff");
    }

    #[test]
    fn test_parse_tiff_big_endian_with_resolution() {
        let blob = tiff_bytes(false, 64, 32, Some((96, 96, 2)));
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.px_width, 64);
        assert_eq!(info.px_height, 32);
        assert_eq!(info.horz_dpi, 96);
        assert_eq!(info.vert_dpi, 96);
    }

    #[test]
    fn test_parse_tiff_resolution_unit_centimeter() {
        // unit 3 = per cm; 100 px/cm * 2.54 = 254 dpi
        let blob = tiff_bytes(true, 10, 10, Some((100, 100, 3)));
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.horz_dpi, 254);
        assert_eq!(info.vert_dpi, 254);
    }

    /// TIFF with dimensions stored as SHORT (type 3, value inline in the first
    /// 2 bytes of the value field). Real-world big-endian TIFFs often do this;
    /// reading the value as u32 would shift it up by 16 bits.
    fn tiff_bytes_short_dims(little: bool, w: u16, h: u16) -> Vec<u8> {
        let mut v = Vec::new();
        v.extend_from_slice(if little { b"II" } else { b"MM" });
        push16(&mut v, little, 42);
        push32(&mut v, little, 8); // IFD offset
        push16(&mut v, little, 2); // entry count
        for (tag, dim) in [(256u16, w), (257u16, h)] {
            push16(&mut v, little, tag);
            push16(&mut v, little, 3); // SHORT
            push32(&mut v, little, 1);
            push16(&mut v, little, dim);
            push16(&mut v, little, 0); // padding
        }
        push32(&mut v, little, 0); // next IFD offset
        v
    }

    #[test]
    fn test_parse_tiff_short_dimensions_both_endians() {
        let info = ImageInfo::parse(&tiff_bytes_short_dims(true, 300, 200)).unwrap();
        assert_eq!(info.px_width, 300);
        assert_eq!(info.px_height, 200);
        let info = ImageInfo::parse(&tiff_bytes_short_dims(false, 300, 200)).unwrap();
        assert_eq!(info.px_width, 300);
        assert_eq!(info.px_height, 200);
    }

    #[test]
    fn test_parse_tiff_no_resolution_tags_defaults_72() {
        let blob = tiff_bytes(true, 100, 100, None);
        let info = ImageInfo::parse(&blob).unwrap();
        assert_eq!(info.horz_dpi, 72);
        assert_eq!(info.vert_dpi, 72);
    }

    #[test]
    fn test_parse_tiff_missing_dimensions_errors() {
        // valid header, IFD with zero entries -> width/height stay 0
        let mut v = Vec::new();
        v.extend_from_slice(b"II");
        v.extend_from_slice(&42u16.to_le_bytes());
        v.extend_from_slice(&8u32.to_le_bytes());
        v.extend_from_slice(&0u16.to_le_bytes()); // 0 IFD entries
        v.extend_from_slice(&0u32.to_le_bytes());
        assert_eq!(
            ImageInfo::parse(&v).unwrap_err(),
            "invalid TIFF: missing dimensions"
        );
    }

    #[test]
    fn test_parse_tiff_ifd_offset_out_of_bounds_errors() {
        let mut v = Vec::new();
        v.extend_from_slice(b"MM\x00*");
        v.extend_from_slice(&9999u32.to_be_bytes()); // IFD offset beyond EOF
        assert_eq!(ImageInfo::parse(&v).unwrap_err(), "invalid TIFF");
    }

    // ---------- EMU conversion ----------

    #[test]
    fn test_width_emu_exact_inch() {
        // 100 px at 100 dpi = exactly 1 inch = 914400 EMU
        assert_eq!(info(100, 50, 100, 100).width_emu(), 914400);
    }

    #[test]
    fn test_height_emu_rounds_to_nearest() {
        // 914400 / 7 = 130628.57... -> 130629
        assert_eq!(info(1, 1, 7, 7).height_emu(), 130629);
    }

    #[test]
    fn test_width_emu_uses_horizontal_dpi_only() {
        let i = info(72, 50, 72, 999);
        assert_eq!(i.width_emu(), 914400);
        assert_ne!(i.height_emu(), 914400); // 50px at 999dpi
    }

    // ---------- scaled_dimensions ----------

    #[test]
    fn test_scaled_dimensions_native_when_none() {
        let i = info(100, 50, 100, 100);
        assert_eq!(i.scaled_dimensions(None, None), (914400, 457200));
    }

    #[test]
    fn test_scaled_dimensions_both_given_verbatim() {
        let i = info(100, 50, 100, 100);
        assert_eq!(i.scaled_dimensions(Some(1), Some(2)), (1, 2));
    }

    #[test]
    fn test_scaled_dimensions_width_only_preserves_aspect() {
        let i = info(100, 50, 100, 100); // native 914400 x 457200
        assert_eq!(i.scaled_dimensions(Some(1828800), None), (1828800, 914400));
    }

    #[test]
    fn test_scaled_dimensions_height_only_preserves_aspect() {
        let i = info(100, 50, 100, 100);
        assert_eq!(i.scaled_dimensions(None, Some(228600)), (457200, 228600));
    }

    // ---------- length helpers ----------

    #[test]
    fn test_length_unit_constants() {
        use super::length::*;
        assert_eq!(inches(1.0), 914400);
        assert_eq!(cm(1.0), 360000);
        assert_eq!(mm(1.0), 36000);
        assert_eq!(pt(1.0), 12700);
        assert_eq!(twips(1.0), 635);
    }

    #[test]
    fn test_length_conversions_consistent_with_inches() {
        use super::length::*;
        assert_eq!(cm(2.54), inches(1.0));
        assert_eq!(pt(72.0), inches(1.0));
        assert_eq!(twips(1440.0), inches(1.0));
        assert_eq!(inches(1.5), 1371600);
    }
}
