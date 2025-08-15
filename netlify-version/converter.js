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
                
                // Store the markdown text for conversion
                this.currentMarkdownText = text;
                
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
        
        // Update convert button to show "Convert Another File"
        const convertBtn = document.getElementById('convertBtn');
        convertBtn.disabled = false;
        convertBtn.textContent = 'Convert Another File';
        convertBtn.classList.add('btn-secondary');
        convertBtn.classList.remove('btn-primary');
        
        // Auto-reset after a short delay to make it ready for next conversion
        setTimeout(() => {
            this.resetForm();
            convertBtn.textContent = 'Convert Files';
            convertBtn.classList.add('btn-primary');
            convertBtn.classList.remove('btn-secondary');
        }, 3000);
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
        zip.file('word/numbering.xml', DocxStructure.getNumberingXml());
        zip.file('word/document.xml', DocxStructure.getDocumentXml(docContent));
        
        // Generate blob
        return await zip.generateAsync({
            type: "blob", 
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        });
    }

    convertHtmlToWordML(html) {
        // Enhanced HTML to WordML conversion that preserves formatting
        
        // Step 1: Parse the markdown tokens directly for better control
        const tokens = marked.lexer(this.currentMarkdownText || '');
        return this.processTokensToWordML(tokens);
    }

    processTokensToWordML(tokens) {
        let wordML = '';
        
        for (const token of tokens) {
            switch (token.type) {
                case 'heading':
                    const headingLevel = Math.min(token.depth, 6); // Limit to H6
                    const headingStyle = `Heading${headingLevel}`;
                    wordML += `<w:p><w:pPr><w:pStyle w:val="${headingStyle}"/></w:pPr><w:r><w:t>${this.escapeXml(token.text)}</w:t></w:r></w:p>`;
                    break;
                    
                case 'paragraph':
                    wordML += this.processParagraphToken(token);
                    break;
                    
                case 'list':
                    wordML += this.processListToken(token);
                    break;
                    
                case 'blockquote':
                    wordML += `<w:p><w:pPr><w:pStyle w:val="Quote"/><w:ind w:left="720"/></w:pPr><w:r><w:t>${this.escapeXml(token.text)}</w:t></w:r></w:p>`;
                    break;
                    
                case 'code':
                    // Code block
                    const codeLines = token.text.split('\n');
                    for (const line of codeLines) {
                        wordML += `<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">${this.escapeXml(line)}</w:t></w:r></w:p>`;
                    }
                    break;
                    
                case 'space':
                    // Add line break for spaces
                    wordML += '<w:p><w:r><w:t></w:t></w:r></w:p>';
                    break;
                    
                case 'hr':
                    // Horizontal rule
                    wordML += '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr><w:r><w:t></w:t></w:r></w:p>';
                    break;
                    
                default:
                    if (token.text) {
                        wordML += `<w:p><w:r><w:t>${this.escapeXml(token.text)}</w:t></w:r></w:p>`;
                    }
            }
        }
        
        return wordML || '<w:p><w:r><w:t></w:t></w:r></w:p>'; // Ensure at least one paragraph
    }

    processParagraphToken(token) {
        // Process paragraph with inline formatting
        if (token.tokens && token.tokens.length > 0) {
            let runs = '';
            for (const inlineToken of token.tokens) {
                runs += this.processInlineToken(inlineToken);
            }
            return `<w:p><w:pPr></w:pPr>${runs}</w:p>`;
        } else if (token.text) {
            // Parse the raw text for inline markdown if tokens aren't available
            return `<w:p><w:pPr></w:pPr>${this.parseInlineMarkdown(token.text)}</w:p>`;
        } else {
            // Empty paragraph
            return `<w:p><w:r><w:t></w:t></w:r></w:p>`;
        }
    }

    processInlineToken(token) {
        switch (token.type) {
            case 'text':
                return `<w:r><w:t>${this.escapeXml(token.text)}</w:t></w:r>`;
                
            case 'strong':
                return `<w:r><w:rPr><w:b/></w:rPr><w:t>${this.escapeXml(token.text)}</w:t></w:r>`;
                
            case 'em':
                return `<w:r><w:rPr><w:i/></w:rPr><w:t>${this.escapeXml(token.text)}</w:t></w:r>`;
                
            case 'codespan':
                return `<w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/></w:rPr><w:t>${this.escapeXml(token.text)}</w:t></w:r>`;
                
            case 'del':
                return `<w:r><w:rPr><w:strike/></w:rPr><w:t>${this.escapeXml(token.text)}</w:t></w:r>`;
                
            case 'link':
                return `<w:r><w:rPr><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr><w:t>${this.escapeXml(token.text)}</w:t></w:r>`;
                
            case 'br':
                return '<w:r><w:br/></w:r>';
                
            default:
                return `<w:r><w:t>${this.escapeXml(token.text || '')}</w:t></w:r>`;
        }
    }

    processListToken(token) {
        let wordML = '';
        const isOrdered = token.ordered;
        const numId = isOrdered ? 1 : 2; // 1 for ordered, 2 for unordered
        
        for (let i = 0; i < token.items.length; i++) {
            const item = token.items[i];
            let itemRuns = '';
            
            if (item.tokens) {
                for (const itemToken of item.tokens) {
                    if (itemToken.type === 'text') {
                        itemRuns += this.parseInlineMarkdown(itemToken.text);
                    } else if (itemToken.type === 'paragraph') {
                        if (itemToken.tokens) {
                            for (const inlineToken of itemToken.tokens) {
                                itemRuns += this.processInlineToken(inlineToken);
                            }
                        } else if (itemToken.text) {
                            itemRuns += this.parseInlineMarkdown(itemToken.text);
                        }
                    }
                }
            } else if (item.text) {
                itemRuns = this.parseInlineMarkdown(item.text);
            }
            
            // Ensure we have at least an empty run
            if (!itemRuns) {
                itemRuns = '<w:r><w:t></w:t></w:r>';
            }
            
            wordML += `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>${itemRuns}</w:p>`;
        }
        
        return wordML;
    }

    parseInlineMarkdown(text) {
        if (!text) return '<w:r><w:t></w:t></w:r>';
        
        // Split text into segments, processing markdown formatting
        const segments = this.parseMarkdownSegments(text);
        let result = '';
        
        for (const segment of segments) {
            result += this.createRun(segment.text, segment.formatting);
        }
        
        return result || '<w:r><w:t></w:t></w:r>';
    }

    parseMarkdownSegments(text) {
        const segments = [];
        let currentIndex = 0;
        
        // Define patterns for different markdown syntax
        const patterns = [
            { regex: /\*\*\*(.+?)\*\*\*/g, formatting: { bold: true, italic: true } }, // Bold + Italic
            { regex: /\*\*(.+?)\*\*/g, formatting: { bold: true } },                    // Bold
            { regex: /\*(.+?)\*/g, formatting: { italic: true } },                      // Italic
            { regex: /___(.+?)___/g, formatting: { bold: true, italic: true } },        // Bold + Italic (alt)
            { regex: /__(.+?)__/g, formatting: { bold: true } },                        // Bold (alt)
            { regex: /_(.+?)_/g, formatting: { italic: true } },                        // Italic (alt)
            { regex: /~~(.+?)~~/g, formatting: { strikethrough: true } },               // Strikethrough
            { regex: /`(.+?)`/g, formatting: { code: true } },                          // Inline code
            { regex: /\[([^\]]+)\]\(([^)]+)\)/g, formatting: { link: true } }           // Links
        ];
        
        // Find all matches
        const matches = [];
        for (const pattern of patterns) {
            let match;
            pattern.regex.lastIndex = 0; // Reset regex
            while ((match = pattern.regex.exec(text)) !== null) {
                matches.push({
                    start: match.index,
                    end: match.index + match[0].length,
                    text: match[1], // Captured group (content without markdown syntax)
                    fullMatch: match[0],
                    formatting: pattern.formatting,
                    linkUrl: pattern.formatting.link ? match[2] : null
                });
            }
        }
        
        // Sort matches by position
        matches.sort((a, b) => a.start - b.start);
        
        // Remove overlapping matches (keep the first one)
        const filteredMatches = [];
        for (const match of matches) {
            const overlaps = filteredMatches.some(existing => 
                (match.start < existing.end && match.end > existing.start)
            );
            if (!overlaps) {
                filteredMatches.push(match);
            }
        }
        
        // Build segments
        let lastEnd = 0;
        for (const match of filteredMatches) {
            // Add plain text before this match
            if (match.start > lastEnd) {
                const plainText = text.substring(lastEnd, match.start);
                if (plainText) {
                    segments.push({ text: plainText, formatting: {} });
                }
            }
            
            // Add formatted text
            segments.push({ 
                text: match.text, 
                formatting: match.formatting,
                linkUrl: match.linkUrl
            });
            
            lastEnd = match.end;
        }
        
        // Add remaining plain text
        if (lastEnd < text.length) {
            const remainingText = text.substring(lastEnd);
            if (remainingText) {
                segments.push({ text: remainingText, formatting: {} });
            }
        }
        
        // If no matches found, return the entire text as plain
        if (segments.length === 0) {
            segments.push({ text: text, formatting: {} });
        }
        
        return segments;
    }

    createRun(text, formatting = {}) {
        if (!text) return '';
        
        let rPr = '';
        
        // Apply formatting
        if (formatting.bold) {
            rPr += '<w:b/>';
        }
        if (formatting.italic) {
            rPr += '<w:i/>';
        }
        if (formatting.strikethrough) {
            rPr += '<w:strike/>';
        }
        if (formatting.code) {
            rPr += '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>';
            rPr += '<w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/>';
        }
        if (formatting.link) {
            rPr += '<w:color w:val="0000FF"/>';
            rPr += '<w:u w:val="single"/>';
        }
        
        const rPrTag = rPr ? `<w:rPr>${rPr}</w:rPr>` : '';
        const escapedText = this.escapeXml(text);
        
        return `<w:r>${rPrTag}<w:t>${escapedText}</w:t></w:r>`;
    }

    escapeXml(text) {
        if (!text) return '';
        return text.replace(/&/g, '&amp;')
                  .replace(/</g, '&lt;')
                  .replace(/>/g, '&gt;')
                  .replace(/"/g, '&quot;')
                  .replace(/'/g, '&apos;');
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
