#!/usr/bin/env python3
"""
SUPER EASY SETUP SCRIPT
📋 This script will set up everything you need to convert markdown to DOCX!

Just double-click this file and follow the instructions.
No technical knowledge required!
"""

import subprocess
import sys
import os
from pathlib import Path

def print_header():
    print("=" * 70)
    print("🎉 WELCOME TO THE MARKDOWN TO DOCX CONVERTER SETUP! 🎉")
    print("=" * 70)
    print()
    print("This setup will install everything you need to convert your")
    print("markdown (.md) files to Word (.docx) format!")
    print()

def check_python():
    print("🐍 Checking Python version...")
    version = sys.version_info
    if version.major >= 3 and version.minor >= 6:
        print(f"✅ Python {version.major}.{version.minor}.{version.micro} - Perfect!")
        return True
    else:
        print(f"❌ Python {version.major}.{version.minor}.{version.micro} is too old.")
        print("Please install Python 3.6 or newer from https://python.org")
        return False

def install_dependencies():
    print("\n📦 Installing Python dependencies...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✅ Python dependencies installed successfully!")
        return True
    except subprocess.CalledProcessError:
        print("❌ Failed to install Python dependencies.")
        print("You might need to run this as administrator or use a virtual environment.")
        return False
    except FileNotFoundError:
        print("❌ requirements.txt file not found!")
        return False

def check_pandoc():
    print("\n📄 Checking for Pandoc...")
    try:
        subprocess.check_output(['pandoc', '--version'], stderr=subprocess.STDOUT)
        print("✅ Pandoc is already installed!")
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("⚠️  Pandoc is not installed.")
        return False

def show_pandoc_instructions():
    print("\n📋 PANDOC INSTALLATION INSTRUCTIONS:")
    print("-" * 40)
    
    if sys.platform == "darwin":  # macOS
        print("For macOS:")
        print("1. Install Homebrew if you don't have it: https://brew.sh")
        print("2. Run: brew install pandoc")
        print("3. Or download from: https://pandoc.org/installing.html")
    elif sys.platform == "win32":  # Windows
        print("For Windows:")
        print("1. Download the installer from: https://pandoc.org/installing.html")
        print("2. Run the installer and follow the instructions")
        print("3. Restart your command prompt after installation")
    else:  # Linux
        print("For Linux:")
        print("Ubuntu/Debian: sudo apt-get install pandoc")
        print("CentOS/RHEL: sudo yum install pandoc")
        print("Or download from: https://pandoc.org/installing.html")

def test_conversion():
    print("\n🧪 Testing the converter...")
    
    # Create a test markdown file
    test_content = """# Test Document

This is a **test** to make sure everything works!

## Features
- Conversion works ✅
- Your setup is complete! 🎉

*Happy converting!*
"""
    
    test_file = Path("setup_test.md")
    try:
        with open(test_file, "w") as f:
            f.write(test_content)
        
        # Try to convert it
        try:
            import pypandoc
            pypandoc.convert_file(str(test_file), 'docx', outputfile='setup_test.docx')
            
            if Path("setup_test.docx").exists():
                print("✅ Test conversion successful!")
                print("   Created: setup_test.docx")
                
                # Clean up test files
                test_file.unlink(missing_ok=True)
                Path("setup_test.docx").unlink(missing_ok=True)
                return True
            else:
                print("❌ Test conversion failed - no output file created")
                return False
                
        except Exception as e:
            print(f"❌ Test conversion failed: {e}")
            return False
            
    except Exception as e:
        print(f"❌ Could not create test file: {e}")
        return False
    finally:
        # Clean up test file if it exists
        test_file.unlink(missing_ok=True)

def show_usage_instructions():
    print("\n" + "=" * 70)
    print("🎉 SETUP COMPLETE! 🎉")
    print("=" * 70)
    print()
    print("📝 HOW TO USE YOUR CONVERTER:")
    print()
    print("OPTION 1 - Simple Interface (Recommended):")
    print("   Double-click: convert_drag_drop.py")
    print("   Follow the prompts!")
    print()
    print("OPTION 2 - Web Interface (Super Fancy!):")
    print("   Double-click: markdown_converter_web.py")
    print("   Open your browser to: http://localhost:5000")
    print()
    print("OPTION 3 - Command Line (For Experts):")
    print("   python3 markdown_to_docx_converter.py -i your_folder")
    print()
    print("💡 TIP: Put your .md files in a folder, then use any option above!")
    print()
    print("🆘 Need help? Check the README.md file!")
    print("=" * 70)

def main():
    print_header()
    
    # Check Python version
    if not check_python():
        input("\nPress Enter to exit...")
        return
    
    # Install Python dependencies
    print("\n" + "-" * 50)
    if not install_dependencies():
        input("\nPress Enter to exit...")
        return
    
    # Check for Pandoc
    print("\n" + "-" * 50)
    pandoc_installed = check_pandoc()
    
    if not pandoc_installed:
        show_pandoc_instructions()
        print("\n⚠️  Please install Pandoc and run this setup again.")
        input("\nPress Enter to exit...")
        return
    
    # Test everything works
    print("\n" + "-" * 50)
    if test_conversion():
        show_usage_instructions()
    else:
        print("\n❌ Setup completed but there might be issues.")
        print("Please check that all dependencies are properly installed.")
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 Setup cancelled. Goodbye!")
    except Exception as e:
        print(f"\n❌ An unexpected error occurred: {e}")
        input("Press Enter to exit...")
