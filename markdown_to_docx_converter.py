#!/usr/bin/env python3
"""
Markdown to DOCX Converter

A Python script that converts all markdown (.md) files in a specified directory
and its subdirectories to Microsoft Word (.docx) format while preserving the
original folder structure.

Features:
- Batch conversion of multiple markdown files
- Preserves folder structure in output
- Timestamped output directories to avoid overwrites
- Detailed progress reporting
- Error handling with graceful continuation

Author: Brennan Kenneth Brown
License: MIT
Repository: https://github.com/[your-username]/markdown-to-docx-converter
"""

import os
import sys
import argparse
from pathlib import Path
import pypandoc
from datetime import datetime

def setup_output_directory(source_dir, output_base=None):
    """Create output directory structure"""
    if output_base:
        output_dir = Path(output_base)
    else:
        output_dir = source_dir.parent / f"{source_dir.name}_DOCX_Export_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir

def convert_markdown_to_docx(md_file_path, output_path):
    """Convert a single markdown file to docx"""
    try:
        # Use pypandoc to convert markdown to docx
        pypandoc.convert_file(
            str(md_file_path), 
            'docx', 
            outputfile=str(output_path),
            extra_args=[
                '--standalone'
            ]
        )
        return True
    except Exception as e:
        print(f"Error converting {md_file_path}: {str(e)}")
        return False

def find_markdown_files(directory):
    """Recursively find all markdown files in directory"""
    markdown_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.md'):
                markdown_files.append(Path(root) / file)
    return markdown_files

def preserve_folder_structure(source_file, source_root, output_root):
    """Create the same folder structure in output directory"""
    relative_path = source_file.parent.relative_to(source_root)
    output_folder = output_root / relative_path
    output_folder.mkdir(parents=True, exist_ok=True)
    return output_folder

def main():
    """Main conversion function"""
    parser = argparse.ArgumentParser(
        description="Convert markdown files to DOCX format while preserving folder structure",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s                           # Convert files in current directory
  %(prog)s -i docs/                  # Convert files in docs/ directory
  %(prog)s -i notes/ -o converted/   # Specify input and output directories
        """
    )
    
    parser.add_argument(
        '-i', '--input', 
        type=str, 
        default='.',
        help='Input directory containing markdown files (default: current directory)'
    )
    
    parser.add_argument(
        '-o', '--output',
        type=str,
        help='Output directory for converted files (default: auto-generated timestamped folder)'
    )
    
    parser.add_argument(
        '--version',
        action='version',
        version='%(prog)s 1.0.0'
    )
    
    args = parser.parse_args()
    
    # Define the source directory
    source_dir = Path(args.input).resolve()
    
    if not source_dir.exists():
        print(f"Error: Input directory '{source_dir}' not found!")
        sys.exit(1)
    
    if not source_dir.is_dir():
        print(f"Error: '{source_dir}' is not a directory!")
        sys.exit(1)
    
    # Create output directory
    output_dir = setup_output_directory(source_dir, args.output)
    print(f"Converting markdown files from: {source_dir}")
    print(f"Output directory: {output_dir}")
    print("-" * 60)
    
    # Find all markdown files
    markdown_files = find_markdown_files(source_dir)
    
    if not markdown_files:
        print("No markdown files found in the directory.")
        return
    
    print(f"Found {len(markdown_files)} markdown files to convert:")
    
    # Convert each file
    successful_conversions = 0
    failed_conversions = 0
    
    for md_file in markdown_files:
        # Preserve folder structure
        output_folder = preserve_folder_structure(md_file, source_dir, output_dir)
        
        # Create output filename
        docx_filename = md_file.stem + ".docx"
        output_path = output_folder / docx_filename
        
        print(f"Converting: {md_file.relative_to(source_dir)} -> {output_path.relative_to(output_dir)}")
        
        # Convert the file
        if convert_markdown_to_docx(md_file, output_path):
            successful_conversions += 1
            print(f"  ✓ Success")
        else:
            failed_conversions += 1
            print(f"  ✗ Failed")
    
    print("-" * 60)
    print(f"Conversion complete!")
    print(f"Successfully converted: {successful_conversions} files")
    print(f"Failed conversions: {failed_conversions} files")
    print(f"Output location: {output_dir}")
    
    if failed_conversions > 0:
        print("\nNote: Some files failed to convert. This might be due to:")
        print("- Complex markdown syntax not supported by pandoc")
        print("- File encoding issues")
        print("- Missing pandoc installation")

def check_dependencies():
    """Check if required dependencies are available"""
    try:
        import pypandoc
        # Check if pandoc is installed
        pypandoc.get_pandoc_version()
        return True
    except ImportError:
        print("Error: pypandoc is not installed.")
        print("Please install it with: pip install pypandoc")
        return False
    except OSError:
        print("Error: pandoc is not installed on your system.")
        print("Please install pandoc:")
        print("  macOS: brew install pandoc")
        print("  Linux: sudo apt-get install pandoc  # or equivalent for your distro")
        print("  Windows: Download from https://pandoc.org/installing.html")
        return False

if __name__ == "__main__":
    if not check_dependencies():
        sys.exit(1)
    
    main()
