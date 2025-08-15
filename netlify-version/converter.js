/**
 * Markdown to DOCX Converter - JavaScript Module
 * Handles file upload, conversion, and download functionality
 */

class MarkdownConverter {
    constructor() {
        this.selectedFiles = [];
        this.initializeEventListeners();
    }

    initializeEventListeners() {
        // File input change handler
        document.getElementById('fileInput').addEventListener('change', (e) => {
            this.selectedFiles = Array.from(e.target.files);
            this.updateUploadArea();
        });

        // Drag and drop handlers
        const uploadArea = document.querySelector('.upload-area');
        
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        
        uploadArea.addEventListener('dragleave', (e) => {
            uploadArea.classList.remove('dragover');
        });
        
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            
            const files = Array.from(e.dataTransfer.files).filter(file => file.name.endsWith('.md'));
            if (files.length > 0) {
                this.selectedFiles = files;
                this.updateUploadArea();
            }
        });

        // Upload area click handler
        uploadArea.addEventListener('click', () => {
            document.getElementById('fileInput').click();
        });
    }

    updateUploadArea() {
        const uploadArea = document.querySelector('.upload-area');
        const convertBtn = document.getElementById('convertBtn');
        
        if (this.selectedFiles.length > 0) {
            uploadArea.innerHTML = `
                <div class="text-6xl mb-4">✅</div>
                <div class="text-xl font-semibold text-gray-800 mb-2">${this.selectedFiles.length} markdown file(s) selected</div>
                <div class="text-gray-600">Ready to convert!</div>
            `;
            convertBtn.disabled = false;
        } else {
            uploadArea.innerHTML = `
                <div class="text-6xl mb-4 text-blue-500">☁️</div>
                <div class="text-xl font-semibold text-gray-800 mb-2">Click to select markdown files</div>
                <div class="text-gray-600">or drag and drop your .md files here</div>
            `;
            convertBtn.disabled = true;
        }
    }

    async startConversion() {
        if (this.selectedFiles.length === 0) return;
        
        // Show progress
        document.getElementById('progressContainer').classList.remove('hidden');
        document.getElementById('convertBtn').disabled = true;
        document.getElementById('results').classList.add('hidden');
        document.getElementById('downloadSection').classList.add('hidden');
        
        const fileResults = document.getElementById('fileResults');
        fileResults.innerHTML = '';
        
        let successCount = 0;
        let failCount = 0;
        
        for (let i = 0; i < this.selectedFiles.length; i++) {
            const file = this.selectedFiles[i];
            const progress = ((i + 1) / this.selectedFiles.length) * 100;
            
            document.getElementById('progressFill').style.width = progress + '%';
            document.getElementById('progressText').textContent = `Converting ${file.name}...`;
            
            try {
                // Read file content
                const text = await this.readFileAsText(file);
                
                // Convert markdown to DOCX
                const docxBlob = await this.convertMarkdownToDocx(text, file.name);
                
                // Create download link
                const fileName = file.name.replace('.md', '.docx');
                this.downloadFile(docxBlob, fileName);
                
                // Add to results
                const resultDiv = document.createElement('div');
                resultDiv.className = 'flex justify-between items-center py-2 border-b border-gray-200 last:border-b-0';
                resultDiv.innerHTML = `
                    <span class="text-gray-800">${file.name}</span>
                    <span class="success font-semibold">✅ Converted</span>
                `;
                fileResults.appendChild(resultDiv);
                
                successCount++;
            } catch (error) {
                console.error('Conversion error:', error);
                
                const resultDiv = document.createElement('div');
                resultDiv.className = 'flex justify-between items-center py-2 border-b border-gray-200 last:border-b-0';
                resultDiv.innerHTML = `
                    <span class="text-gray-800">${file.name}</span>
                    <span class="error font-semibold">❌ Failed</span>
                `;
                fileResults.appendChild(resultDiv);
                
                failCount++;
            }
            
            // Small delay for UI updates
            await new Promise(resolve => setTimeout(resolve, 100));
        }
        
        // Show results
        document.getElementById('results').classList.remove('hidden');
        document.getElementById('downloadSection').classList.remove('hidden');
        document.getElementById('progressText').textContent = `Completed! ${successCount} files converted successfully`;
        document.getElementById('convertBtn').disabled = false;
    }

    readFileAsText(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = e => resolve(e.target.result);
            reader.onerror = reject;
            reader.readAsText(file);
        });
    }

    async convertMarkdownToDocx(markdownText, fileName) {
        // Convert markdown to HTML
        const html = marked.parse(markdownText);
        
        // Convert HTML to simple Word XML
        const docContent = this.convertHtmlToWordML(html);
        
        // Create DOCX package using JSZip
        const zip = new JSZip();
        
        // Add required DOCX structure files
        zip.file('[Content_Types].xml', DocxStructure.getContentTypesXml());
        zip.file('_rels/.rels', DocxStructure.getRelsXml());
        zip.file('word/_rels/document.xml.rels', DocxStructure.getDocumentRelsXml());
        zip.file('word/styles.xml', DocxStructure.getStylesXml());
        zip.file('word/document.xml', DocxStructure.getDocumentXml(docContent));
        
        // Generate blob
        return await zip.generateAsync({
            type: "blob", 
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        });
    }

    convertHtmlToWordML(html) {
        // Simple HTML to WordML conversion
        let wordML = html;
        
        // Convert basic HTML tags to WordML
        wordML = wordML.replace(/<h1[^>]*>(.*?)<\/h1>/gi, '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>$1</w:t></w:r></w:p>');
        wordML = wordML.replace(/<h2[^>]*>(.*?)<\/h2>/gi, '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>$1</w:t></w:r></w:p>');
        wordML = wordML.replace(/<h3[^>]*>(.*?)<\/h3>/gi, '<w:p><w:pPr><w:pStyle w:val="Heading3"/></w:pPr><w:r><w:t>$1</w:t></w:r></w:p>');
        wordML = wordML.replace(/<p[^>]*>(.*?)<\/p>/gi, '<w:p><w:r><w:t>$1</w:t></w:r></w:p>');
        wordML = wordML.replace(/<strong[^>]*>(.*?)<\/strong>/gi, '<w:r><w:rPr><w:b/></w:rPr><w:t>$1</w:t></w:r>');
        wordML = wordML.replace(/<b[^>]*>(.*?)<\/b>/gi, '<w:r><w:rPr><w:b/></w:rPr><w:t>$1</w:t></w:r>');
        wordML = wordML.replace(/<em[^>]*>(.*?)<\/em>/gi, '<w:r><w:rPr><w:i/></w:rPr><w:t>$1</w:t></w:r>');
        wordML = wordML.replace(/<i[^>]*>(.*?)<\/i>/gi, '<w:r><w:rPr><w:i/></w:rPr><w:t>$1</w:t></w:r>');
        wordML = wordML.replace(/<br\s*\/?>/gi, '<w:br/>');
        
        // Remove remaining HTML tags and decode entities
        wordML = wordML.replace(/<[^>]+>/g, '');
        wordML = wordML.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"');
        
        // Wrap in paragraph if no paragraphs found
        if (!wordML.includes('<w:p>')) {
            wordML = `<w:p><w:r><w:t>${wordML}</w:t></w:r></w:p>`;
        }
        
        return wordML;
    }

    downloadFile(blob, fileName) {
        saveAs(blob, fileName);
    }

    resetForm() {
        this.selectedFiles = [];
        document.getElementById('fileInput').value = '';
        document.getElementById('convertBtn').disabled = true;
        document.getElementById('progressContainer').classList.add('hidden');
        document.getElementById('results').classList.add('hidden');
        document.getElementById('downloadSection').classList.add('hidden');
        
        const uploadArea = document.querySelector('.upload-area');
        uploadArea.innerHTML = `
            <div class="text-6xl mb-4 text-blue-500">☁️</div>
            <div class="text-xl font-semibold text-gray-800 mb-2">Click to select markdown files</div>
            <div class="text-gray-600">or drag and drop your .md files here</div>
        `;
    }
}

// Initialize converter when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    window.converter = new MarkdownConverter();
});

// Global functions for button handlers
function startConversion() {
    window.converter.startConversion();
}

function resetForm() {
    window.converter.resetForm();
}
