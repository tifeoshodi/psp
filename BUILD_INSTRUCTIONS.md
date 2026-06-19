# Building Standalone Executable

This guide explains how to create a standalone executable of the Project Scheduler that doesn't require Python to be installed on end-user machines.

## Quick Start (Recommended)

### Option 1: Automated Build Script
1. **Run the build script:**
   ```bash
   python build_executable.py
   ```
   
2. **Or use the batch file (Windows):**
   ```bash
   build.bat
   ```

The executable will be created in the `dist/` folder as `ProjectScheduler.exe`.

## Manual Build Process

### Prerequisites
1. **Install PyInstaller:**
   ```bash
   pip install -r requirements.txt
   ```

### Build Commands

#### Basic Build (Single File)
```bash
pyinstaller --onefile --windowed --name ProjectScheduler project_scheduler.py
```

#### Advanced Build (With Logo and Optimizations)
```bash
pyinstaller --onefile --windowed --name ProjectScheduler --icon IESL-Logo.png --add-data "IESL-Logo.png;." --clean project_scheduler.py
```

#### Using Spec File (Recommended)
```bash
pyinstaller project_scheduler.spec
```

## Build Options Explained

| Option | Description |
|--------|-------------|
| `--onefile` | Creates a single executable file |
| `--windowed` | Hides the console window (GUI app) |
| `--name` | Sets the executable name |
| `--icon` | Sets the executable icon |
| `--add-data` | Includes data files (logo) |
| `--clean` | Cleans build cache before building |

## Output Structure

After building, you'll get:
```
dist/
└── ProjectScheduler.exe    # Standalone executable (~15-25 MB)

build/                      # Temporary build files (can be deleted)
ProjectScheduler.spec       # Build configuration file
```

## Distribution

### What to Distribute
- **Single file:** `dist/ProjectScheduler.exe`
- **Size:** Approximately 15-25 MB
- **Dependencies:** All included (no Python required)

### System Requirements for End Users
- **OS:** Windows 7 or later
- **Architecture:** 64-bit (default) or 32-bit
- **RAM:** Minimum 512 MB
- **Disk:** ~50 MB free space

## Troubleshooting

### Common Issues

#### 1. "Module not found" errors
**Solution:** Add missing modules to `hiddenimports` in the spec file:
```python
hiddenimports=['missing_module_name']
```

#### 2. Logo not showing in executable
**Solution:** Ensure `IESL-Logo.png` exists and is included:
```bash
--add-data "IESL-Logo.png;."
```

#### 3. Executable too large
**Solution:** Exclude unused packages:
```python
excludes=['matplotlib', 'numpy', 'pandas']
```

#### 4. Antivirus false positive
**Solution:** 
- Add exception in antivirus software
- Sign the executable (for production)
- Use `--noupx` flag if UPX compression causes issues

### Build Environment Tips

1. **Use virtual environment:**
   ```bash
   python -m venv build_env
   build_env\Scripts\activate
   pip install -r requirements.txt
   ```

2. **Clean builds:**
   ```bash
   pyinstaller --clean project_scheduler.spec
   ```

3. **Test thoroughly:**
   - Test on different Windows versions
   - Test on machines without Python
   - Test all application features

## Advanced Customization

### Custom Icon
Replace `IESL-Logo.png` with your `.ico` file:
```bash
--icon your_icon.ico
```

### Splash Screen
Add loading splash:
```bash
pip install pyi-splash
# Add splash code to main script
```

### Version Information
Create version file for executable properties:
```bash
pyi-makespec --version-file version.txt project_scheduler.py
```

## File Size Optimization

### Reduce Executable Size
1. **Exclude unused modules:**
   ```python
   excludes=['tkinter.dnd', 'unittest', 'pdb']
   ```

2. **Use UPX compression:**
   ```bash
   --upx-dir /path/to/upx
   ```

3. **Directory build instead of onefile:**
   ```bash
   pyinstaller --windowed project_scheduler.py
   ```
   (Creates smaller startup time but multiple files)

## Security Considerations

### Code Signing (Production)
```bash
# Sign the executable (requires certificate)
signtool sign /f certificate.pfx /p password ProjectScheduler.exe
```

### Obfuscation
Consider using PyArmor for source code protection:
```bash
pip install pyarmor
pyarmor obfuscate project_scheduler.py
```

## Performance Notes

- **Startup time:** 2-5 seconds (normal for PyInstaller)
- **Memory usage:** ~50-100 MB (includes Python runtime)
- **File access:** Logo embedded, no external dependencies

## Deployment Checklist

- [ ] Executable runs on clean Windows machine
- [ ] All GUI features work correctly
- [ ] Excel generation functions properly
- [ ] Logo displays correctly
- [ ] File dialogs work
- [ ] Error handling works
- [ ] No console window appears
- [ ] File size acceptable (<50 MB)

## Support

For build issues:
1. Check PyInstaller documentation
2. Test in clean environment
3. Review error logs in `build/` directory
4. Ensure all dependencies are in requirements.txt