# Rust GDI+ Snapshot Tool

A Rust reimplementation of GDI+ screen-capture functionality, inspired by AutoHotkey libraries such as [AHKv2-Gdip](https://github.com/mmikeww/AHKv2-Gdip).

Supports Windows 10+ and exports to PNG, JPG, BMP, GIF, and TIFF formats.

This tool was created as an open-source alternative to closed-source utilities such as nircmd.

## Build
```
cargo build --release
```

## Usage
```
gdip_snapshot output.jpg                # Grab screenshot of primary monitor
gdip_snapshot --full output.jpg         # Capture full virtual desktop (all monitors)
gdip_snapshot 0 0 1920 1080 output.jpg  # Grab 1920x1080 screenshot starting at (0, 0)
```

All Rust source code is original and independently written.  
Licensed under the terms of the [MIT License](LICENSE.md)
