#!/usr/bin/env python3
"""
Build script for creating Windows executable
Run this script to build the Windows executable locally
"""

import os
import sys
import subprocess
import shutil

def run_command(cmd):
    """Run a command and return the result"""
    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"Error: {result.stderr}")
        return False
    print(f"Output: {result.stdout}")
    return True

def build_executable():
    """Build the Windows executable"""
    print("Building Windows executable...")
    
    # Clean previous builds
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    if os.path.exists('build'):
        shutil.rmtree('build')
    
    # Install dependencies
    print("Installing dependencies...")
    if not run_command([sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt']):
        return False
    
    # Build using PyInstaller
    print("Building with PyInstaller...")
    cmd = [
        'pyinstaller',
        '--onefile',
        '--windowed',
        '--name=PAS Wireless Device Exporter',
        '--add-data=main.py;.',
        '--hidden-import=pymodbus',
        '--hidden-import=pymodbus.client',
        '--hidden-import=openpyxl',
        '--hidden-import=tkinter',
        '--hidden-import=tkinter.filedialog',
        '--hidden-import=tkinter.messagebox',
        'main.py'
    ]
    
    # Add icon if it exists
    if os.path.exists('icon.ico'):
        cmd.extend(['--icon=icon.ico'])
    
    if not run_command(cmd):
        return False
    
    print("Build completed successfully!")
    print(f"Executable created: {os.path.abspath('dist/PAS Wireless Device Exporter.exe')}")
    return True

if __name__ == "__main__":
    if build_executable():
        print("Build successful!")
    else:
        print("Build failed!")
        sys.exit(1)
