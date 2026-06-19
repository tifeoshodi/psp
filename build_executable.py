#!/usr/bin/env python3
"""
Script to build standalone executable for Project Scheduler
"""

import os
import subprocess
import sys
from pathlib import Path

def build_executable():
    """Build the standalone executable using PyInstaller"""
    
    print("🚀 Building Project Scheduler executable...")
    print("=" * 50)
    
    # Check if PyInstaller is installed
    try:
        import PyInstaller
        print(f"✅ PyInstaller found: {PyInstaller.__version__}")
    except ImportError:
        print("❌ PyInstaller not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller>=5.0.0"])
        print("✅ PyInstaller installed successfully")
    
    # Define build parameters
    script_name = "project_scheduler.py"
    app_name = "ProjectScheduler"
    
    # Check if main script exists
    if not os.path.exists(script_name):
        print(f"❌ Error: {script_name} not found in current directory")
        return False
    
    # Check if logo exists
    logo_file = "IESL-Logo.png"
    logo_exists = os.path.exists(logo_file)
    if logo_exists:
        print(f"✅ Logo found: {logo_file} - will be embedded in executable")
    else:
        print(f"⚠️  Warning: {logo_file} not found - logo will not be embedded")
        print("   Make sure IESL-Logo.png is in the same directory as project_scheduler.py")
    
    # Build command with all necessary options
    build_cmd = [
        "pyinstaller",
        "--onefile",                    # Create single executable file
        "--windowed",                   # Hide console window (for GUI app)
        "--name", app_name,             # Set executable name
        "--icon", logo_file if os.path.exists(logo_file) else None,  # Use logo as icon
        "--add-data", f"{logo_file};." if logo_exists else "",  # Include logo file
        "--distpath", "dist",           # Output directory
        "--workpath", "build",          # Build directory
        "--specpath", ".",              # Spec file location
        "--clean",                      # Clean cache before building
        script_name
    ]
    
    # Remove None values and empty strings
    build_cmd = [arg for arg in build_cmd if arg]
    
    try:
        print(f"🔨 Running PyInstaller...")
        print(f"Command: {' '.join(build_cmd)}")
        print()
        
        # Run PyInstaller
        result = subprocess.run(build_cmd, check=True, capture_output=True, text=True)
        
        print("✅ Build completed successfully!")
        print()
        
        # Check if executable was created
        exe_path = Path("dist") / f"{app_name}.exe"
        if exe_path.exists():
            exe_size = exe_path.stat().st_size / (1024 * 1024)  # Size in MB
            print(f"📦 Executable created: {exe_path}")
            print(f"📏 File size: {exe_size:.1f} MB")
            print()
            print("🎉 SUCCESS! Your standalone executable is ready!")
            print(f"   Location: {exe_path.absolute()}")
            print()
            print("📋 Distribution Notes:")
            print("   • The executable includes all dependencies")
            print("   • No Python installation required on target machines")
            if logo_exists:
                print("   • Logo file is embedded in the executable")
            else:
                print("   • Logo file was not found - Excel files will generate without logo")
            print("   • Can be distributed as a single .exe file")
            return True
        else:
            print("❌ Error: Executable was not created")
            return False
            
    except subprocess.CalledProcessError as e:
        print(f"❌ Build failed with error:")
        print(f"   {e}")
        if e.stdout:
            print(f"STDOUT: {e.stdout}")
        if e.stderr:
            print(f"STDERR: {e.stderr}")
        return False
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return False

def cleanup_build_files():
    """Clean up build artifacts"""
    import shutil
    
    print("\n🧹 Cleaning up build files...")
    
    # Directories to clean
    cleanup_dirs = ["build", "__pycache__"]
    cleanup_files = ["*.spec"]
    
    for dir_name in cleanup_dirs:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"   Removed: {dir_name}/")
    
    import glob
    for pattern in cleanup_files:
        for file in glob.glob(pattern):
            os.remove(file)
            print(f"   Removed: {file}")

if __name__ == "__main__":
    success = build_executable()
    
    if success:
        # Ask user if they want to clean up build files
        response = input("\n🗑️  Clean up build files? (y/N): ").strip().lower()
        if response in ['y', 'yes']:
            cleanup_build_files()
            print("✅ Cleanup completed")
    
    print("\n" + "=" * 50)
    print("Build process finished.")
    input("Press Enter to exit...")