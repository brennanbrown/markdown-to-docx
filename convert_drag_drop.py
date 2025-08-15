#!/usr/bin/env python3
"""
Markdown to DOCX Converter - Simple Drag & Drop Version

SUPER EASY TO USE:
1. Double-click this file
2. When prompted, enter the path to your folder with markdown files
3. Press Enter and wait for conversion to complete!

No technical knowledge required!
"""

import os
import sys
from pathlib import Path

def simple_user_interface():
    """Simple text-based interface for non-technical users"""
    print("=" * 60)
    print("🎉 WELCOME TO THE MARKDOWN TO DOCX CONVERTER! 🎉")
    print("=" * 60)
    print()
    print("This tool will convert all your .md files to .docx format!")
    print("Your folder structure will be preserved.")
    print()
    
    # Check dependencies first
    print("⏳ Checking if everything is ready...")
    try:
        import pypandoc
        pypandoc.get_pandoc_version()
        print("✅ All dependencies are ready!")
    except ImportError:
        print("❌ Missing dependency: pypandoc")
        print("Please run: pip install pypandoc")
        input("Press Enter to exit...")
        return
    except OSError:
        print("❌ Missing dependency: pandoc")
        print("Please install pandoc from: https://pandoc.org/installing.html")
        input("Press Enter to exit...")
        return
    
    print()
    print("-" * 60)
    
    while True:
        print("📁 STEP 1: Choose your markdown files folder")
        print()
        print("Options:")
        print("1. Type the folder path")
        print("2. Use current folder (if you put this script in your markdown folder)")
        print("3. Exit")
        print()
        
        choice = input("Enter your choice (1, 2, or 3): ").strip()
        
        if choice == "3":
            print("👋 Goodbye!")
            return
        elif choice == "2":
            folder_path = "."
        elif choice == "1":
            print()
            print("📝 Enter the full path to your folder containing .md files:")
            print("Example: /Users/yourname/Documents/my-notes")
            print("Example: C:\\Users\\yourname\\Documents\\my-notes")
            print()
            folder_path = input("Folder path: ").strip().strip('"').strip("'")
        else:
            print("❌ Invalid choice. Please try again.")
            print()
            continue
        
        # Validate folder
        folder_path = Path(folder_path).resolve()
        if not folder_path.exists():
            print(f"❌ Folder not found: {folder_path}")
            print("Please check the path and try again.")
            print()
            continue
        
        if not folder_path.is_dir():
            print(f"❌ This is not a folder: {folder_path}")
            print("Please provide a folder path.")
            print()
            continue
        
        break
    
    print()
    print(f"✅ Using folder: {folder_path}")
    print()
    
    # Ask about output location
    print("📤 STEP 2: Choose output location")
    print("1. Auto-generate a new folder (recommended)")
    print("2. Choose a specific output folder")
    print()
    
    output_choice = input("Enter your choice (1 or 2): ").strip()
    
    if output_choice == "2":
        print()
        output_path = input("Enter output folder path: ").strip().strip('"').strip("'")
        output_path = Path(output_path) if output_path else None
    else:
        output_path = None
    
    print()
    print("🚀 Starting conversion...")
    print("=" * 60)
    
    # Import and run conversion
    try:
        from markdown_to_docx_converter import (
            find_markdown_files, 
            preserve_folder_structure, 
            convert_markdown_to_docx,
            setup_output_directory
        )
        
        # Find markdown files
        markdown_files = find_markdown_files(folder_path)
        
        if not markdown_files:
            print("⚠️  No .md files found in the selected folder.")
            print("Make sure your folder contains markdown files with .md extension.")
            input("Press Enter to exit...")
            return
        
        print(f"📋 Found {len(markdown_files)} markdown files")
        
        # Setup output
        output_dir = setup_output_directory(folder_path, output_path)
        print(f"📁 Output will be saved to: {output_dir}")
        print()
        
        # Convert files
        successful = 0
        failed = 0
        
        for i, md_file in enumerate(markdown_files, 1):
            print(f"🔄 Converting file {i}/{len(markdown_files)}: {md_file.name}")
            
            # Preserve folder structure
            output_folder = preserve_folder_structure(md_file, folder_path, output_dir)
            docx_filename = md_file.stem + ".docx"
            output_path_file = output_folder / docx_filename
            
            if convert_markdown_to_docx(md_file, output_path_file):
                successful += 1
                print(f"   ✅ Success!")
            else:
                failed += 1
                print(f"   ❌ Failed")
        
        print()
        print("=" * 60)
        print("🎉 CONVERSION COMPLETE! 🎉")
        print(f"✅ Successfully converted: {successful} files")
        if failed > 0:
            print(f"❌ Failed: {failed} files")
        print(f"📁 Your converted files are in: {output_dir}")
        print("=" * 60)
        
        # Ask to open folder
        print()
        open_folder = input("Would you like to open the output folder? (y/n): ").strip().lower()
        if open_folder.startswith('y'):
            try:
                if sys.platform == "darwin":  # macOS
                    os.system(f'open "{output_dir}"')
                elif sys.platform == "win32":  # Windows
                    os.system(f'explorer "{output_dir}"')
                else:  # Linux
                    os.system(f'xdg-open "{output_dir}"')
            except:
                print(f"Could not open folder automatically. Please navigate to: {output_dir}")
        
    except Exception as e:
        print(f"❌ An error occurred: {e}")
        print("Please check that all dependencies are installed correctly.")
    
    print()
    input("Press Enter to exit...")

if __name__ == "__main__":
    simple_user_interface()
