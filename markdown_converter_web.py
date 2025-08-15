#!/usr/bin/env python3
"""
Markdown to DOCX Converter - Web Interface

A beautiful, user-friendly web interface for converting markdown files.
No technical knowledge required - just open in your browser!

Features:
- Drag & drop folder selection
- Real-time progress updates
- Beautiful, modern interface
- Works on any device with a web browser

Author: Brennan Kenneth Brown
License: MIT
"""

from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
import os
import sys
import tempfile
import zipfile
import threading
import time
from pathlib import Path
import shutil
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'markdown-converter-secret-key'

# Global variables for progress tracking
conversion_progress = {"status": "idle", "progress": 0, "message": "", "files": [], "output_path": ""}

@app.route('/')
def index():
    """Main page with upload interface"""
    return render_template('index.html')

@app.route('/favicon.ico')
def favicon():
    """Handle favicon requests"""
    return '', 204  # No content

@app.route('/convert', methods=['POST'])
def convert_files():
    """Handle file conversion request"""
    global conversion_progress
    
    # Reset progress
    conversion_progress = {"status": "starting", "progress": 0, "message": "Starting conversion...", "files": [], "output_path": ""}
    
    # Get uploaded files and read their content immediately
    uploaded_files = request.files.getlist('markdown_files')
    
    # Read file contents in the main thread (before Flask closes the file objects)
    file_data = []
    for file in uploaded_files:
        if file and file.filename and file.filename.strip() != '' and file.filename.endswith('.md'):
            try:
                # Read the file content immediately
                content = file.read()
                file_data.append({
                    'filename': file.filename,
                    'content': content
                })
            except Exception as e:
                return jsonify({"error": f"Error reading file {file.filename}: {str(e)}"}), 400
    
    if not file_data:
        return jsonify({"error": "No valid markdown files selected"}), 400
    
    # Start conversion in background thread with file data (not file objects)
    thread = threading.Thread(target=process_conversion, args=(file_data,))
    thread.daemon = True
    thread.start()
    
    return jsonify({"message": "Conversion started", "status": "processing"})

@app.route('/progress')
def get_progress():
    """Get current conversion progress"""
    return jsonify(conversion_progress)

@app.route('/download')
def download_results():
    """Download converted files as ZIP"""
    if conversion_progress.get("output_path") and os.path.exists(conversion_progress["output_path"]):
        return send_file(conversion_progress["output_path"], as_attachment=True, 
                        download_name=f"converted_docx_files_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip")
    else:
        return jsonify({"error": "No files available for download"}), 404

def process_conversion(file_data):
    """Process the uploaded file data and convert them"""
    global conversion_progress
    
    try:
        # Check dependencies
        conversion_progress["message"] = "Checking dependencies..."
        conversion_progress["progress"] = 5
        
        try:
            import pypandoc
            pypandoc.get_pandoc_version()
        except ImportError:
            conversion_progress["status"] = "error"
            conversion_progress["message"] = "Error: pypandoc not installed. Please install with: pip install pypandoc"
            return
        except OSError:
            conversion_progress["status"] = "error"
            conversion_progress["message"] = "Error: pandoc not installed. Please install from https://pandoc.org/"
            return
        
        # Create temporary directories
        temp_input_dir = tempfile.mkdtemp(prefix="md_converter_input_")
        temp_output_dir = tempfile.mkdtemp(prefix="md_converter_output_")
        
        conversion_progress["message"] = "Processing uploaded files..."
        conversion_progress["progress"] = 10
        
        # Save file data to disk
        markdown_files = []
        for file_info in file_data:
            filename = file_info['filename']
            content = file_info['content']
            
            # Ensure safe filename
            safe_filename = filename.replace('/', '_').replace('\\', '_')
            file_path = os.path.join(temp_input_dir, safe_filename)
            
            try:
                # Write content to disk
                with open(file_path, 'wb') as f:
                    f.write(content)
                
                # Verify file was saved
                if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
                    markdown_files.append(Path(file_path))
                else:
                    conversion_progress["status"] = "error"
                    conversion_progress["message"] = f"Failed to save file {filename} - file is empty"
                    return
                    
            except Exception as e:
                conversion_progress["status"] = "error"
                conversion_progress["message"] = f"Error saving file {filename}: {str(e)}"
                return
        
        if not markdown_files:
            conversion_progress["status"] = "error"
            conversion_progress["message"] = "No markdown files found in uploaded files"
            return
        
        conversion_progress["message"] = f"Found {len(markdown_files)} markdown files"
        conversion_progress["progress"] = 20
        
        # Import conversion functions
        from markdown_to_docx_converter import convert_markdown_to_docx
        
        # Convert files
        successful = 0
        failed = 0
        converted_files = []
        
        for i, md_file in enumerate(markdown_files):
            progress = 20 + (i / len(markdown_files)) * 70  # 20% to 90%
            conversion_progress["progress"] = int(progress)
            conversion_progress["message"] = f"Converting {md_file.name}..."
            
            # Create output file path
            docx_filename = md_file.stem + ".docx"
            output_path = Path(temp_output_dir) / docx_filename
            
            try:
                # Ensure the md_file exists and is readable
                if not md_file.exists():
                    failed += 1
                    converted_files.append({"name": md_file.name, "status": "failed", "output": None})
                    continue
                
                # Convert the file
                if convert_markdown_to_docx(md_file, output_path):
                    # Verify output file was created
                    if output_path.exists() and output_path.stat().st_size > 0:
                        successful += 1
                        converted_files.append({"name": md_file.name, "status": "success", "output": docx_filename})
                    else:
                        failed += 1
                        converted_files.append({"name": md_file.name, "status": "failed", "output": None})
                else:
                    failed += 1
                    converted_files.append({"name": md_file.name, "status": "failed", "output": None})
            except Exception as e:
                failed += 1
                converted_files.append({"name": md_file.name, "status": "failed", "output": None})
                print(f"Error converting {md_file.name}: {e}")  # For debugging
        
        conversion_progress["files"] = converted_files
        conversion_progress["progress"] = 90
        conversion_progress["message"] = "Creating download package..."
        
        # Create ZIP file with converted documents
        zip_path = os.path.join(tempfile.gettempdir(), f"converted_files_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip")
        
        try:
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                files_added = 0
                for file_info in converted_files:
                    if file_info["status"] == "success" and file_info["output"]:
                        file_path = Path(temp_output_dir) / file_info["output"]
                        if file_path.exists() and file_path.stat().st_size > 0:
                            zipf.write(file_path, file_info["output"])
                            files_added += 1
                
                if files_added == 0:
                    conversion_progress["status"] = "error"
                    conversion_progress["message"] = "No files were successfully converted"
                    return
                    
        except Exception as e:
            conversion_progress["status"] = "error"
            conversion_progress["message"] = f"Error creating download package: {str(e)}"
            return
        
        # Cleanup temp directories
        shutil.rmtree(temp_input_dir, ignore_errors=True)
        shutil.rmtree(temp_output_dir, ignore_errors=True)
        
        # Final status
        conversion_progress["status"] = "completed"
        conversion_progress["progress"] = 100
        conversion_progress["message"] = f"Conversion complete! {successful} files converted successfully"
        conversion_progress["output_path"] = zip_path
        
        if failed > 0:
            conversion_progress["message"] += f", {failed} files failed"
            
    except Exception as e:
        conversion_progress["status"] = "error"
        conversion_progress["message"] = f"Error during conversion: {str(e)}"

def create_html_template():
    """Create the HTML template for the web interface"""
    template_dir = Path("templates")
    template_dir.mkdir(exist_ok=True)
    
    html_content = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Markdown to DOCX Converter</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 800px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            padding: 40px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
        }
        
        .header {
            text-align: center;
            margin-bottom: 40px;
        }
        
        .header h1 {
            color: #333;
            font-size: 2.5em;
            margin-bottom: 10px;
        }
        
        .header p {
            color: #666;
            font-size: 1.2em;
        }
        
        .upload-area {
            border: 3px dashed #ddd;
            border-radius: 15px;
            padding: 60px 40px;
            text-align: center;
            margin-bottom: 30px;
            transition: all 0.3s ease;
            cursor: pointer;
        }
        
        .upload-area:hover {
            border-color: #667eea;
            background: #f8f9ff;
        }
        
        .upload-area.dragover {
            border-color: #667eea;
            background: #f0f4ff;
        }
        
        .upload-icon {
            font-size: 3em;
            margin-bottom: 20px;
            color: #667eea;
        }
        
        .upload-text {
            font-size: 1.3em;
            color: #333;
            margin-bottom: 10px;
        }
        
        .upload-subtext {
            color: #666;
            font-size: 1em;
        }
        
        .file-input {
            display: none;
        }
        
        .btn {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            padding: 15px 30px;
            border-radius: 50px;
            font-size: 1.1em;
            font-weight: bold;
            cursor: pointer;
            transition: all 0.3s ease;
            margin: 10px;
        }
        
        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(102, 126, 234, 0.3);
        }
        
        .btn:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }
        
        .progress-container {
            display: none;
            margin-top: 30px;
        }
        
        .progress-bar {
            width: 100%;
            height: 20px;
            background: #f0f0f0;
            border-radius: 10px;
            overflow: hidden;
            margin-bottom: 15px;
        }
        
        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
            width: 0%;
            transition: width 0.3s ease;
        }
        
        .progress-text {
            text-align: center;
            color: #333;
            font-weight: bold;
        }
        
        .results {
            display: none;
            margin-top: 30px;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 10px;
        }
        
        .file-result {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 10px 0;
            border-bottom: 1px solid #eee;
        }
        
        .file-result:last-child {
            border-bottom: none;
        }
        
        .success {
            color: #28a745;
        }
        
        .error {
            color: #dc3545;
        }
        
        .download-section {
            text-align: center;
            margin-top: 30px;
            display: none;
        }
        
        @media (max-width: 600px) {
            .container {
                padding: 20px;
            }
            
            .header h1 {
                font-size: 2em;
            }
            
            .upload-area {
                padding: 40px 20px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📄➡️📘 Markdown to DOCX</h1>
            <p>Convert your markdown files to Word documents instantly!</p>
        </div>
        
        <div class="upload-area" onclick="document.getElementById('fileInput').click()">
            <div class="upload-icon">☁️</div>
            <div class="upload-text">Click to select markdown files</div>
            <div class="upload-subtext">or drag and drop your .md files here</div>
        </div>
        
        <form id="uploadForm" enctype="multipart/form-data">
            <input type="file" id="fileInput" name="markdown_files" multiple accept=".md" class="file-input">
        </form>
        
        <div style="text-align: center;">
            <button id="convertBtn" class="btn" onclick="startConversion()" disabled>Convert Files</button>
            <button class="btn" onclick="resetForm()">Reset</button>
        </div>
        
        <div id="progressContainer" class="progress-container">
            <div class="progress-bar">
                <div id="progressFill" class="progress-fill"></div>
            </div>
            <div id="progressText" class="progress-text">Starting conversion...</div>
        </div>
        
        <div id="results" class="results">
            <h3>Conversion Results:</h3>
            <div id="fileResults"></div>
        </div>
        
        <div id="downloadSection" class="download-section">
            <button id="downloadBtn" class="btn" onclick="downloadFiles()">📥 Download Converted Files</button>
        </div>
    </div>

    <script>
        let selectedFiles = [];
        
        // File input change handler
        document.getElementById('fileInput').addEventListener('change', function(e) {
            selectedFiles = Array.from(e.target.files);
            updateUploadArea();
        });
        
        // Drag and drop handlers
        const uploadArea = document.querySelector('.upload-area');
        
        uploadArea.addEventListener('dragover', function(e) {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        
        uploadArea.addEventListener('dragleave', function(e) {
            uploadArea.classList.remove('dragover');
        });
        
        uploadArea.addEventListener('drop', function(e) {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            
            const files = Array.from(e.dataTransfer.files).filter(file => file.name.endsWith('.md'));
            if (files.length > 0) {
                selectedFiles = files;
                updateUploadArea();
            }
        });
        
        function updateUploadArea() {
            const convertBtn = document.getElementById('convertBtn');
            
            if (selectedFiles.length > 0) {
                uploadArea.innerHTML = `
                    <div class="upload-icon">✅</div>
                    <div class="upload-text">${selectedFiles.length} markdown file(s) selected</div>
                    <div class="upload-subtext">Ready to convert!</div>
                `;
                convertBtn.disabled = false;
            } else {
                convertBtn.disabled = true;
            }
        }
        
        function startConversion() {
            const formData = new FormData();
            selectedFiles.forEach(file => {
                formData.append('markdown_files', file);
            });
            
            document.getElementById('progressContainer').style.display = 'block';
            document.getElementById('convertBtn').disabled = true;
            
            fetch('/convert', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                if (data.error) {
                    alert('Error: ' + data.error);
                    resetProgress();
                } else {
                    checkProgress();
                }
            })
            .catch(error => {
                alert('Error: ' + error);
                resetProgress();
            });
        }
        
        function checkProgress() {
            fetch('/progress')
            .then(response => response.json())
            .then(data => {
                updateProgress(data);
                
                if (data.status === 'completed') {
                    showResults(data);
                } else if (data.status === 'error') {
                    alert('Error: ' + data.message);
                    resetProgress();
                } else {
                    setTimeout(checkProgress, 1000);
                }
            });
        }
        
        function updateProgress(data) {
            document.getElementById('progressFill').style.width = data.progress + '%';
            document.getElementById('progressText').textContent = data.message;
        }
        
        function showResults(data) {
            const results = document.getElementById('results');
            const fileResults = document.getElementById('fileResults');
            
            fileResults.innerHTML = '';
            
            data.files.forEach(file => {
                const div = document.createElement('div');
                div.className = 'file-result';
                div.innerHTML = `
                    <span>${file.name}</span>
                    <span class="${file.status}">${file.status === 'success' ? '✅ Converted' : '❌ Failed'}</span>
                `;
                fileResults.appendChild(div);
            });
            
            results.style.display = 'block';
            document.getElementById('downloadSection').style.display = 'block';
        }
        
        function downloadFiles() {
            window.location.href = '/download';
        }
        
        function resetForm() {
            selectedFiles = [];
            document.getElementById('fileInput').value = '';
            document.getElementById('convertBtn').disabled = true;
            document.getElementById('progressContainer').style.display = 'none';
            document.getElementById('results').style.display = 'none';
            document.getElementById('downloadSection').style.display = 'none';
            
            uploadArea.innerHTML = `
                <div class="upload-icon">☁️</div>
                <div class="upload-text">Click to select markdown files</div>
                <div class="upload-subtext">or drag and drop your .md files here</div>
            `;
        }
        
        function resetProgress() {
            document.getElementById('progressContainer').style.display = 'none';
            document.getElementById('convertBtn').disabled = false;
        }
    </script>
</body>
</html>
    '''
    
    with open(template_dir / "index.html", "w") as f:
        f.write(html_content)

def main():
    """Run the web application"""
    print("🌐 Starting Markdown to DOCX Web Converter...")
    
    # Check dependencies
    try:
        import pypandoc
        pypandoc.get_pandoc_version()
        print("✅ Dependencies check passed")
    except ImportError:
        print("❌ Missing dependency: pypandoc")
        print("Please install with: pip install pypandoc")
        return
    except OSError:
        print("❌ Missing dependency: pandoc")
        print("Please install from: https://pandoc.org/installing.html")
        return
    
    # Create HTML template
    create_html_template()
    
    print("🚀 Starting web server...")
    print("📱 Open your browser and go to: http://localhost:8080")
    print("🛑 Press Ctrl+C to stop the server")
    print()
    
    try:
        app.run(host='0.0.0.0', port=8080, debug=False)
    except KeyboardInterrupt:
        print("\n👋 Server stopped. Goodbye!")

if __name__ == "__main__":
    main()
