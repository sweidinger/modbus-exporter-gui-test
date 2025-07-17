#!/usr/bin/env python3
"""
Setup script for Modbus Data Exporter GUI Application
"""

import re
import os
from setuptools import setup, find_packages

# Read version from main.py
def get_version():
    """Extract version from main.py"""
    with open("main.py", "r") as f:
        content = f.read()
        match = re.search(r'__version__ = "([^"]+)"', content)
        if match:
            return match.group(1)
        else:
            raise RuntimeError("Cannot find version string in main.py")

# Read the README file
def get_long_description():
    """Get long description from README.md if it exists"""
    if os.path.exists("README.md"):
        with open("README.md", "r", encoding="utf-8") as f:
            return f.read()
    return "Modbus Data Exporter GUI Application"

setup(
    name="modbus-exporter-gui",
    version=get_version(),
    author="Stefan Weidinger",
    author_email="stefan@example.com",
    description="A GUI application for exporting Modbus device data to CSV or Excel format",
    long_description=get_long_description(),
    long_description_content_type="text/markdown",
    url="https://github.com/sweidinger/modbus-exporter-gui-test",
    py_modules=["main"],
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Software Development :: Libraries",
        "Topic :: System :: Hardware",
        "Topic :: System :: Monitoring",
    ],
    python_requires=">=3.7",
    install_requires=[
        "pymodbus>=3.0.0",
        "openpyxl>=3.0.0",
    ],
    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "black>=22.0.0",
            "flake8>=4.0.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "modbus-exporter-gui=main:main",
        ],
    },
    keywords="modbus, data export, gui, excel, csv, diagnostics",
    project_urls={
        "Bug Reports": "https://github.com/sweidinger/modbus-exporter-gui-test/issues",
        "Source": "https://github.com/sweidinger/modbus-exporter-gui-test",
    },
)
