# PDF Matcher - EXE Build Guide

## Quick Start

### Step 1: Build EXE File

Double-click to run:
```
build_exe_final.bat
```

This will automatically:
- Install PyInstaller (if needed)
- Clean previous builds
- Create EXE file

**Time required:** 3-5 minutes

### Step 2: Check Generated File

After build completes, find the file at:

```
dist/
  └── PDF정렬프로그램.exe  ← Use this file!
```

---

## Distribution

### Method 1: Single File (Recommended)

1. Copy `dist\PDF정렬프로그램.exe` only
2. Paste to desired folder
3. Double-click to run!

**Benefits:**
- Just one file to copy
- No Python installation needed
- Works on Windows 10/11 immediately

---

## Build Options

### Modify spec file

Open `PDFMatcher.spec` and edit:

```python
# Add icon
icon='icon.ico',

# Show console (for debugging)
console=True,

# Include additional data files
datas=[
    ('config.json', '.'),
    ('README.md', '.'),
],
```

---

## Troubleshooting

### Build Fails

**Error: "pyinstaller not found"**
```bash
pip install pyinstaller
```

**Error: "module not found"**
```bash
pip install -r requirements.txt
```

**Error: exe crashes on run**
- Set `console=True` in spec file
- Rebuild and check error messages

### File Size Too Large (100MB+)

This is normal:
- Qt libraries included
- PDF processing libraries included
- Everything in single file

To reduce size:
```bash
# UPX compression (already in spec)
upx=True
```

---

## System Requirements

### Development (for building)
- Windows 10/11
- Python 3.8+
- Required packages installed

### Runtime (for users)
- Windows 10/11
- No Python needed!
- No admin rights needed

---

## Advanced Usage

### Optimized Build

```bash
# Smaller size
pyinstaller --onefile --windowed --strip main.py

# With debug info
pyinstaller --onefile --debug all main.py

# Include specific modules
pyinstaller --onefile --hidden-import=module_name main.py
```

### Distribution to Multiple PCs

1. **USB Drive**
   - Copy `PDF정렬프로그램.exe`
   - Paste on each PC

2. **Network Share**
   - Save exe to shared folder
   - Users run directly

3. **Email Attachment**
   - Check file size (usually 50-100MB)
   - Compress before sending

---

## Checklist

Before build:
- [ ] All tests pass
- [ ] config.json exists
- [ ] requirements.txt up to date
- [ ] main.py no errors

After build:
- [ ] exe file created
- [ ] Run test
- [ ] Check functionality
- [ ] Test on other PC

---

## Help

If problems persist:
1. Delete `build` folder
2. Delete `dist` folder
3. Run `build_exe_final.bat` again

Still not working:
- Reinstall Python
- Reinstall dependencies
- Recreate spec file

---

## References

- PyInstaller docs: https://pyinstaller.org/
- PySide6 docs: https://doc.qt.io/qtforpython/
- Project README: README.md

