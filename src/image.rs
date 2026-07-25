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
        let val = rd_u32(e + 8);
        match tag {
            256 => width = val,
            257 => height = val,
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
