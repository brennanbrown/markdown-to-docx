# Markdown to DOCX Converter

A Python script that batch converts markdown files to Microsoft Word format while preserving folder structure and formatting.

## ✨ Features

- **Batch conversion**: Convert multiple markdown files at once
- **Folder structure preservation**: Maintains your original directory organization
- **Timestamped outputs**: Prevents overwriting previous conversions
- **Cross-platform**: Works on macOS, Linux, and Windows
- **Command-line interface**: Easy to use with flexible options
- **Error handling**: Graceful handling of conversion errors
- **Progress tracking**: Real-time feedback during conversion

## 🚀 Quick Start

### 🎯 SUPER EASY - No Technical Knowledge Required!

**STEP 1: One-Click Setup**
1. Download this project
2. Double-click `EASY_SETUP.py` 
3. Follow the simple instructions
4. Done! 🎉

**STEP 2: Choose Your Favorite Way to Convert**

#### 🖱️ Option A: Simple Click & Convert (Recommended)
- Double-click `convert_drag_drop.py`
- Enter your folder path when asked
- Watch the magic happen!

#### 🌐 Option B: Beautiful Web Interface  
- Double-click `markdown_converter_web.py`
- Open your browser to `http://localhost:5000`
- Drag & drop files in the browser!

#### 💻 Option C: Command Line (For Tech-Savvy Users)
```bash
python3 markdown_to_docx_converter.py -i your_folder
```

### 🔧 Manual Installation (If Needed)

If the easy setup doesn't work, here's the manual way:

1. **Install Python 3.6+** from [python.org](https://python.org)
2. **Install Pandoc:**
   - **macOS:** `brew install pandoc`
   - **Windows:** Download from [pandoc.org](https://pandoc.org/installing.html)
   - **Linux:** `sudo apt-get install pandoc`
3. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

## 📁 Example

Before conversion:
```
my_notes/
├── README.md
├── project_notes/
│   ├── meeting_notes.md
│   └── requirements.md
└── personal/
    └── journal.md
```

After conversion:
```
my_notes_DOCX_Export_20241220_143022/
├── README.docx
├── project_notes/
│   ├── meeting_notes.docx
│   └── requirements.docx
└── personal/
    └── journal.docx
```

## 🛠 Command Line Options

| Option | Description | Default |
|--------|-------------|---------|
| `-i`, `--input` | Input directory containing markdown files | Current directory |
| `-o`, `--output` | Output directory for converted files | Auto-generated timestamped folder |
| `--version` | Show version information | - |
| `-h`, `--help` | Show help message | - |

## 🔧 Advanced Usage

### Virtual Environment (Recommended)

For better dependency management:

```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
python3 markdown_to_docx_converter.py
```

### Automation

Create a shell script for repeated use:

```bash
#!/bin/bash
# convert_notes.sh
cd /path/to/markdown-to-docx-converter
source venv/bin/activate
python3 markdown_to_docx_converter.py -i ~/Documents/notes
```

## 📋 Supported Markdown Features

The converter preserves most standard markdown formatting:

- **Headers** (H1-H6)
- **Bold** and *italic* text
- `Code blocks` and inline code
- Lists (ordered and unordered)
- Links and images
- Tables
- Blockquotes
- Horizontal rules

## ⚠️ Troubleshooting

### Common Issues

**"Pandoc not found" error:**
- Ensure pandoc is installed and in your PATH
- Try running `pandoc --version` to verify installation

**"pypandoc not installed" error:**
- Run `pip install pypandoc`
- If using virtual environment, ensure it's activated

**Permission errors:**
- Check that you have write permissions to the output directory
- On Unix systems, you might need to use `sudo` for system-wide installations

**Conversion failures:**
- Some complex markdown syntax might not convert perfectly
- Check the error messages for specific file issues
- Verify your markdown files are valid

### Getting Help

If you encounter issues:

1. Check the error messages - they often contain helpful information
2. Verify all dependencies are properly installed
3. Test with a simple markdown file first
4. Open an issue on GitHub with details about your problem

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

1. **Fork** the repository
2. **Create** a feature branch (`git checkout -b feature/amazing-feature`)
3. **Commit** your changes (`git commit -m 'Add amazing feature'`)
4. **Push** to the branch (`git push origin feature/amazing-feature`)
5. **Open** a Pull Request

### Development Setup

```bash
git clone https://github.com/[your-username]/markdown-to-docx-converter.git
cd markdown-to-docx-converter
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- [Pandoc](https://pandoc.org/) - The universal document converter
- [pypandoc](https://github.com/JessicaTegner/pypandoc) - Python wrapper for pandoc

## 📊 Project Stats

- **Language**: Python 3.6+
- **Dependencies**: pypandoc, pathlib
- **Platform**: Cross-platform (macOS, Linux, Windows)
- **License**: MIT

---

**Made with ❤️ for the markdown community**

*If this tool helped you, please consider giving it a ⭐ on GitHub!*
