#!/usr/bin/env python3
"""
Build script for Final Whisper
Creates a standalone EXE with bundled dependencies (except Whisper/PyTorch)
"""

import subprocess
import sys
import os
from pathlib import Path

def install_build_dependencies():
    """Install PyInstaller and other build dependencies"""
    print(" Installing build dependencies...")
    
    deps = ["pyinstaller", "sv-ttk", "Pillow"]
    for dep in deps:
        print(f"  Installing {dep}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", dep, "--quiet"])
    
    print("OK - Build dependencies installed\n")

def create_spec_file():
    """Create a PyInstaller spec file for more control"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

import sv_ttk
import os

# Find sv_ttk theme files
sv_ttk_path = os.path.dirname(sv_ttk.__file__)

a = Analysis(
    ['whisper_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Include sv_ttk theme files
        (sv_ttk_path, 'sv_ttk'),
        # Include icon and logo files
        ('icon.ico', '.'),
        ('logo.png', '.'),
        # Include version file
        ('version.txt', '.'),
    ],
    hiddenimports=[
        'sv_ttk',
        'tkinter',
        'tkinter.ttk',
        'tkinter.scrolledtext',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'encodings',
        'encodings.utf_8',
        'encodings.cp1252',
        'encodings.mbcs',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Exclude large ML packages - user must have these installed separately
        # NOTE: Do NOT exclude numpy - it's needed for basic operations
        'whisper',
        'torch',
        'torchaudio', 
        'torchvision',
        'scipy',
        'pandas',
        'matplotlib',
        'cv2',
        'tensorflow',
        'keras',
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FinalWhisper',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)
'''
    
    with open('FinalWhisper.spec', 'w') as f:
        f.write(spec_content)

    print("OK - Created FinalWhisper.spec\n")

def build_exe():
    """Build the EXE using PyInstaller"""
    print(" Building EXE...")
    print("   This may take a few minutes...\n")
    
    result = subprocess.run([
        sys.executable, "-m", "PyInstaller",
        "FinalWhisper.spec",
        "--clean",
        "--noconfirm"
    ], capture_output=False)
    
    if result.returncode == 0:
        exe_path = Path("dist/FinalWhisper.exe")
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"\nOK - Build successful!")
            print(f"   Output: {exe_path.absolute()}")
            print(f"   Size: {size_mb:.1f} MB")
        else:
            print("\nOK - Build completed - check dist/ folder")
    else:
        print("\nERROR - Build failed!")
        return False
    
    return True

def create_readme():
    """Create a README for the distribution"""
    readme = '''# Final Whisper

## Requirements

Before running the EXE, you need to install Whisper and its dependencies:

```bash
pip install openai-whisper
```

For GPU acceleration (recommended), install PyTorch with CUDA:
```bash
pip install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cu118
```

## Usage

1. Double-click `FinalWhisper.exe`
2. Select your video file
3. Choose output folder
4. Click "Start Transcription"

## Features

- Automatic speech recognition using OpenAI Whisper
- Smart subtitle formatting with sentence boundary detection
- Multi-language support (optimized for English and Danish)
- GPU acceleration support
- Optional AI proofreading with Claude
- Professional subtitle output (40 chars/line, 2 lines max)

## Troubleshooting

If the app doesn't start:
1. Make sure Whisper is installed: `pip install openai-whisper`
2. Check if antivirus is blocking the EXE
3. Try running from command line to see errors

For GPU issues:
- The app will detect your NVIDIA GPU automatically
- You can install GPU support using the "Setup GPU" button
- Or manually: pip install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cu118

Created by Final Film
'''
    
    with open('dist/README.txt', 'w') as f:
        f.write(readme)
    
    print("OK - Created dist/README.txt")

def main():
    print("="*60)
    print("  Final Whisper - EXE Builder")
    print("="*60 + "\n")
    
    # Check if whisper_gui.py exists
    if not Path('whisper_gui.py').exists():
        print("ERROR - Error: whisper_gui.py not found in current directory")
        print("   Please run this script from the same folder as whisper_gui.py")
        return 1
    
    try:
        install_build_dependencies()
        create_spec_file()
        
        if build_exe():
            # Create dist folder readme
            Path("dist").mkdir(exist_ok=True)
            create_readme()
            
            print("\n" + "="*60)
            print("  Build Complete!")
            print("="*60)
            print("\nThe EXE is in the 'dist' folder: FinalWhisper.exe")
            print("Users will need to install Whisper separately:")
            print("  pip install openai-whisper")
            print("\nFor GPU support:")
            print("  pip install torch --index-url https://download.pytorch.org/whl/cu118")
        
    except Exception as e:
        print(f"\nERROR - Build error: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
