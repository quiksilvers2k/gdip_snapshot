use std::env;
use std::ffi::{OsStr, c_void};
use std::iter::once;
use std::mem::{size_of, zeroed};
use std::os::windows::ffi::OsStrExt;
use std::ptr::{null, null_mut};

use windows::Win32::Foundation::{E_FAIL, E_INVALIDARG};
use windows::Win32::Graphics::Gdi;
use windows::Win32::Graphics::Gdi::{CAPTUREBLT, ROP_CODE, SRCCOPY};
use windows::Win32::Graphics::GdiPlus;
use windows::Win32::System::Com::{CoTaskMemAlloc, CoTaskMemFree};
use windows::Win32::UI::WindowsAndMessaging::{
    GetSystemMetrics, SM_CXSCREEN, SM_CXVIRTUALSCREEN, SM_CYSCREEN, SM_CYVIRTUALSCREEN,
    SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN,
};
use windows::core::{Error, GUID, HRESULT, PCWSTR};

fn wide<S: AsRef<OsStr>>(s: S) -> Vec<u16> {
    s.as_ref().encode_wide().chain(once(0)).collect()
}

struct EncodersGuard(*mut c_void);
impl Drop for EncodersGuard {
    fn drop(&mut self) {
        unsafe { CoTaskMemFree(Some(self.0)) }
    }
}

struct ScreenDcGuard(Gdi::HDC);
impl Drop for ScreenDcGuard {
    fn drop(&mut self) {
        unsafe {
            Gdi::ReleaseDC(None, self.0);
        }
    }
}

struct DcGuard(Gdi::HDC);
impl Drop for DcGuard {
    fn drop(&mut self) {
        unsafe {
            let _ = Gdi::DeleteDC(self.0);
        }
    }
}

struct BitmapGuard(Gdi::HBITMAP);
impl Drop for BitmapGuard {
    fn drop(&mut self) {
        unsafe {
            let _ = Gdi::DeleteObject(self.0.into());
        }
    }
}

struct SelectGuard {
    dc: Gdi::HDC,
    old: Gdi::HGDIOBJ,
}
impl Drop for SelectGuard {
    fn drop(&mut self) {
        unsafe {
            Gdi::SelectObject(self.dc, self.old);
        }
    }
}

struct GdiplusGuard(usize);
impl GdiplusGuard {
    fn new() -> windows::core::Result<Self> {
        gdip_startup().map(Self)
    }
}
impl Drop for GdiplusGuard {
    fn drop(&mut self) {
        gdip_shutdown(self.0);
    }
}

struct ImgGuard(*mut GdiPlus::GpImage);
impl Drop for ImgGuard {
    fn drop(&mut self) {
        if !self.0.is_null() {
            unsafe { GdiPlus::GdipDisposeImage(self.0) };
        }
    }
}

// find a matching image encoder for an extension (like Gdip_SaveBitmapToFile does).
fn clsid_for_extension(ext: &str) -> windows::core::Result<GUID> {
    let mut num = 0u32;
    let mut size = 0u32;
    unsafe {
        if GdiPlus::GdipGetImageEncodersSize(&mut num, &mut size) != GdiPlus::Ok {
            return Err(Error::new(
                HRESULT(E_FAIL.0),
                "GdipGetImageEncodersSize failed",
            ));
        }
    }
    if num == 0 || size == 0 {
        return Err(Error::new(HRESULT(E_FAIL.0), "No image encoders available"));
    }
    // aligned allocation
    let encoders_ptr = unsafe { CoTaskMemAlloc(size as usize) } as *mut GdiPlus::ImageCodecInfo;
    if encoders_ptr.is_null() {
        return Err(Error::new(HRESULT(E_FAIL.0), "CoTaskMemAlloc failed"));
    }
    // ensure free on all paths
    let _encoders_guard = EncodersGuard(encoders_ptr as *mut c_void);
    unsafe {
        if GdiPlus::GdipGetImageEncoders(num, size, encoders_ptr) != GdiPlus::Ok {
            return Err(Error::new(HRESULT(E_FAIL.0), "GdipGetImageEncoders failed"));
        }
    }
    // normalize the requested extension (".png", ".jpg", ...)
    let want = format!(".{}", ext.trim_start_matches('.')).to_ascii_lowercase();
    // iterate the array portion at the beginning of the buffer. Each struct's pointer
    // fields point into the same 'buf', so 'buf' must stay alive until we finish.
    for i in 0..(num as usize) {
        let info = unsafe { &*encoders_ptr.add(i) };
        // some codecs may not provide FilenameExtension.
        if info.FilenameExtension.is_null() {
            continue;
        }
        // read the UTF-16 NUL-terminated string.
        let p = PCWSTR::from_raw(info.FilenameExtension.0);
        let exts = unsafe { p.to_string()? };
        // patterns look like "*.JPG;*.JPEG;*.JPE;*.JFIF".
        for pat in exts.split(';') {
            let pat = pat.trim().trim_start_matches('*').to_ascii_lowercase(); // ".jpg"
            if pat == want {
                return Ok(info.Clsid);
            }
        }
    }
    Err(Error::new(
        HRESULT(E_FAIL.0),
        "No encoder found for the given extension",
    ))
}

fn gdip_startup() -> windows::core::Result<usize> {
    unsafe {
        let mut input: GdiPlus::GdiplusStartupInput = zeroed();
        input.GdiplusVersion = 1;
        let mut token: usize = 0;
        if GdiPlus::GdiplusStartup(
            &mut token,
            &input,
            null_mut::<GdiPlus::GdiplusStartupOutput>(),
        ) != GdiPlus::Ok
        {
            return Err(Error::new(HRESULT(E_FAIL.0), "GdiplusStartup failed"));
        }
        Ok(token)
    }
}

fn gdip_shutdown(token: usize) {
    unsafe { GdiPlus::GdiplusShutdown(token) };
}

fn make_dib_section(
    w: i32,
    h: i32,
    hdc_palette: Gdi::HDC,
) -> windows::core::Result<(Gdi::HBITMAP, *mut u8)> {
    // 32bpp, bottom-up bitmap (positive height)
    let mut bmi: Gdi::BITMAPINFO = unsafe { zeroed() };
    bmi.bmiHeader.biSize = size_of::<Gdi::BITMAPINFOHEADER>() as u32;
    bmi.bmiHeader.biWidth = w;
    bmi.bmiHeader.biHeight = h; // positive => bottom-up
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 32;
    bmi.bmiHeader.biCompression = Gdi::BI_RGB.0;
    let mut bits: *mut core::ffi::c_void = null_mut();
    // unwrap the Result<HBITMAP> here
    let hbmp: Gdi::HBITMAP = unsafe {
        Gdi::CreateDIBSection(
            Some(hdc_palette),
            &bmi,
            Gdi::DIB_RGB_COLORS,
            &mut bits,
            None, // no file mapping
            0,
        )?
    };
    Ok((hbmp, bits as *mut u8))
}

fn capture_region(x: i32, y: i32, w: i32, h: i32) -> windows::core::Result<Gdi::HBITMAP> {
    let raster_op: ROP_CODE = SRCCOPY | CAPTUREBLT;
    unsafe {
        let hdc_screen = Gdi::GetDC(None);
        if hdc_screen.0.is_null() {
            return Err(Error::new(HRESULT(E_FAIL.0), "GetDC failed"));
        }
        let _screen_guard = ScreenDcGuard(hdc_screen);

        let mem_dc = Gdi::CreateCompatibleDC(Some(hdc_screen));
        if mem_dc.0.is_null() {
            return Err(Error::new(HRESULT(E_FAIL.0), "CreateCompatibleDC failed"));
        }
        let _mem_guard = DcGuard(mem_dc);

        // create target bitmap (deleted automatically unless we forget it)
        let (hbmp, _bits) = make_dib_section(w, h, hdc_screen)?;
        let hbmp_guard = BitmapGuard(hbmp);

        // select it into mem DC; selection restored automatically
        let old = Gdi::SelectObject(mem_dc, hbmp.into());
        if old.is_invalid() {
            return Err(Error::new(HRESULT(E_FAIL.0), "SelectObject failed"));
        }
        let _sel_guard = SelectGuard { dc: mem_dc, old };

        // BitBlt from screen into our DIB
        Gdi::BitBlt(mem_dc, 0, 0, w, h, Some(hdc_screen), x, y, raster_op)?;

        // success: transfer ownership to caller (prevent guard from deleting it)
        std::mem::forget(hbmp_guard);
        Ok(hbmp)
    }
}

// wrap HBITMAP -> GDI+ Bitmap, choose encoder by extension, save
fn save_hbitmap_with_gdiplus(hbmp: Gdi::HBITMAP, filename: &str) -> windows::core::Result<()> {
    let mut bmp: *mut GdiPlus::GpBitmap = null_mut();
    unsafe {
        if GdiPlus::GdipCreateBitmapFromHBITMAP(hbmp, Gdi::HPALETTE(std::ptr::null_mut()), &mut bmp)
            != GdiPlus::Ok
        {
            return Err(Error::new(
                HRESULT(E_FAIL.0),
                "GdipCreateBitmapFromHBITMAP failed",
            ));
        }
    }
    // ensure dispose on all paths
    let _guard = ImgGuard(bmp as *mut GdiPlus::GpImage);
    // Pick encoder by extension.
    let ext = std::path::Path::new(filename)
        .extension()
        .and_then(|e| e.to_str())
        .ok_or_else(|| Error::new(HRESULT(E_INVALIDARG.0), "filename has no extension"))?;
    let clsid = clsid_for_extension(ext)?;
    //save output file
    let wname = wide(filename);
    unsafe {
        if GdiPlus::GdipSaveImageToFile(
            bmp as *mut GdiPlus::GpImage,
            PCWSTR(wname.as_ptr()),
            &clsid,
            null(),
        ) != GdiPlus::Ok
        {
            return Err(Error::new(HRESULT(E_FAIL.0), "GdipSaveImageToFile failed"));
        }
    }
    Ok(())
}

fn usage() {
    eprintln!("Usage:");
    eprintln!("  gdip_snapshot <x> <y> <width> <height> <output_file>");
    eprintln!("  gdip_snapshot --full <output_file>     # all monitors (virtual desktop)");
    eprintln!("  gdip_snapshot --primary <output_file>  # primary monitor only");
    eprintln!("  gdip_snapshot <output_file>            # default: --primary");
}

/// Returns (x, y, w, h) for the chosen screen mode.
fn screen_rect(mode: ScreenMode) -> (i32, i32, i32, i32) {
    match mode {
        ScreenMode::Virtual => {
            // entire virtual desktop (spans all monitors; x/y can be negative)
            let x = unsafe { GetSystemMetrics(SM_XVIRTUALSCREEN) };
            let y = unsafe { GetSystemMetrics(SM_YVIRTUALSCREEN) };
            let w = unsafe { GetSystemMetrics(SM_CXVIRTUALSCREEN) };
            let h = unsafe { GetSystemMetrics(SM_CYVIRTUALSCREEN) };
            (x, y, w, h)
        }
        ScreenMode::Primary => {
            // primary monitor only (origin at 0,0)
            let w = unsafe { GetSystemMetrics(SM_CXSCREEN) };
            let h = unsafe { GetSystemMetrics(SM_CYSCREEN) };
            (0, 0, w, h)
        }
    }
}

#[derive(Clone, Copy)]
enum ScreenMode {
    Virtual,
    Primary,
}

fn capture_rectangle(x: i32, y: i32, w: i32, h: i32, filename: &str) -> windows::core::Result<()> {
    let _gdip = GdiplusGuard::new()?; // starts and shuts down GDI+ automatically
    let hbmp = capture_region(x, y, w, h)?;
    let result = save_hbitmap_with_gdiplus(hbmp, filename);
    unsafe {
        let _ = Gdi::DeleteObject(hbmp.into());
    }
    result
}

fn main() -> windows::core::Result<()> {
    let args: Vec<String> = env::args().collect();
    // Modes:
    // 6 args: x y w h filename
    // 3 args: flag + filename
    // 2 args: filename => --primary
    if args.len() == 6 {
        // explicit rectangle
        let x: i32 = args[1].parse().unwrap_or_else(|_| {
            eprintln!("x must be an integer");
            std::process::exit(1);
        });
        let y: i32 = args[2].parse().unwrap_or_else(|_| {
            eprintln!("y must be an integer");
            std::process::exit(1);
        });
        let w: i32 = args[3].parse().unwrap_or_else(|_| {
            eprintln!("width must be an integer");
            std::process::exit(1);
        });
        let h: i32 = args[4].parse().unwrap_or_else(|_| {
            eprintln!("height must be an integer");
            std::process::exit(1);
        });
        let filename = &args[5];
        if w <= 0 || h <= 0 {
            eprintln!("width and height must be > 0");
            std::process::exit(1);
        }
        capture_rectangle(x, y, w, h, filename)?;
        return Ok(());
    }
    // flag + filename OR just filename
    let (mode, filename) = match args.len() {
        3 => {
            let flag = args[1].as_str();
            let fname = args[2].as_str();
            match flag {
                "--full" => (ScreenMode::Virtual, fname),
                "--primary" => (ScreenMode::Primary, fname),
                _ => {
                    usage();
                    std::process::exit(1);
                }
            }
        }
        2 => (ScreenMode::Primary, args[1].as_str()), // default to primary
        _ => {
            usage();
            std::process::exit(1);
        }
    };
    let (x, y, w, h) = screen_rect(mode);
    if w <= 0 || h <= 0 {
        eprintln!("Detected non-positive screen size: {}x{}", w, h);
        std::process::exit(1);
    }
    capture_rectangle(x, y, w, h, filename)?;
    Ok(())
}
