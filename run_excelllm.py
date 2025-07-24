#!/usr/bin/env python3
"""
Simple launcher for ExcelLLM Web Interface
"""

import subprocess
import sys
import os
from pathlib import Path

def check_dependencies():
    """Check and install required dependencies"""
    required = {
        'fastapi': 'fastapi',
        'uvicorn': 'uvicorn[standard]',
        'openpyxl': 'openpyxl',
        'python-multipart': 'python-multipart'
    }

    missing = []

    for module, package in required.items():
        try:
            __import__(module)
        except ImportError:
            missing.append(package)

    if missing:
        print("Installing required dependencies...")
        subprocess.check_call([
            sys.executable, '-m', 'pip', 'install'] + missing
        )
        print("✓ Dependencies installed successfully!")

def main():
    print("""
    ╔═══════════════════════════════════════════════╗
    ║          ExcelLLM Interactive Launcher        ║
    ╚═══════════════════════════════════════════════╝
    """)

    # Check dependencies
    check_dependencies()

    # Import after dependencies are installed
    try:
        from excelllm_webapp import run_server
    except ImportError:
        print("Error: Could not import excelllm_webapp.")
        print("Make sure excelllm_webapp.py is in the same directory.")
        sys.exit(1)

    # Run server
    try:
        run_server()
    except KeyboardInterrupt:
        print("\n\n✓ Server stopped.")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
