# Markdown to DOCX Converter - Netlify Version 🌐

A **100% client-side** markdown to DOCX converter that runs entirely in the browser. Perfect for Netlify deployment!

## ✨ Features

- 🔒 **100% Private** - Files never leave your browser
- ⚡ **Instant Conversion** - No server processing needed
- 💻 **No Software Required** - Works in any modern browser
- 📱 **Mobile Friendly** - Responsive design for all devices
- 🆓 **Completely Free** - No limits, no registration
- 🌐 **Works Offline** - Once loaded, works without internet

## 🚀 Live Demo

**Visit:** [Your Netlify URL will go here]

## 🔧 How It Works

This version uses pure JavaScript to convert markdown files:

1. **Marked.js** - Converts markdown to HTML
2. **Custom WordML Generator** - Converts HTML to Microsoft Word XML format
3. **PizZip** - Creates the DOCX file structure
4. **FileSaver.js** - Triggers file downloads

## 📦 Deploy to Netlify

### Option 1: Drag & Drop (Easiest)

1. Download the `netlify-version` folder
2. Go to [netlify.com](https://netlify.com)
3. Drag the folder to the deploy area
4. Your site is live! 🎉

### Option 2: GitHub Integration

1. **Push to GitHub:**
   ```bash
   git add netlify-version/
   git commit -m "Add Netlify version"
   git push origin main
   ```

2. **Connect to Netlify:**
   - Go to [netlify.com](https://netlify.com)
   - Click "New site from Git"
   - Connect your GitHub repository
   - Set publish directory to `netlify-version`
   - Deploy!

### Option 3: Netlify CLI

```bash
# Install Netlify CLI
npm install -g netlify-cli

# Deploy from the netlify-version folder
cd netlify-version
netlify deploy

# Deploy to production
netlify deploy --prod
```

## 🎯 Supported Markdown Features

✅ **Fully Supported:**
- Headers (H1-H6)
- **Bold** and *italic* text
- Lists (ordered & unordered)
- Links
- Basic paragraphs
- Line breaks

⚠️ **Limited Support:**
- Code blocks (converted to plain text)
- Tables (basic conversion)
- Images (links only)

❌ **Not Supported:**
- Complex nested structures
- Advanced formatting
- Mathematical expressions

## 🔧 Customization

### Change Colors
Edit the CSS variables in the `<style>` section:

```css
:root {
    --primary-gradient: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    --accent-color: #667eea;
}
```

### Add Features
The conversion logic is in the `convertMarkdownToDocx()` function. You can extend it to support more markdown features.

### Branding
Replace the header content with your own branding:

```html
<h1>Your Brand - Markdown to DOCX</h1>
<p>Your custom description here</p>
```

## 🐛 Troubleshooting

**Files won't convert:**
- Ensure files have `.md` extension
- Check browser console for errors
- Try with a simple markdown file first

**Download not starting:**
- Check if browser is blocking downloads
- Ensure JavaScript is enabled
- Try a different browser

**Formatting looks wrong:**
- This is a basic converter - complex markdown may not convert perfectly
- For advanced features, use the Python version

## 🆚 Comparison with Python Version

| Feature | Netlify Version | Python Version |
|---------|----------------|----------------|
| Privacy | 100% client-side | Server required |
| Setup | Just open in browser | Requires Python + pandoc |
| Speed | Instant | Very fast |
| Markdown Support | Basic | Complete (via pandoc) |
| File Size Limits | Browser memory | No limits |
| Offline Use | Yes (after loading) | Yes |

## 📈 Analytics (Optional)

Add Google Analytics or other tracking:

```html
<!-- Add before closing </head> tag -->
<script async src="https://www.googletagmanager.com/gtag/js?id=GA_MEASUREMENT_ID"></script>
<script>
  window.dataLayer = window.dataLayer || [];
  function gtag(){dataLayer.push(arguments);}
  gtag('js', new Date());
  gtag('config', 'GA_MEASUREMENT_ID');
</script>
```

## 🤝 Contributing

Want to improve the converter? Here are some ideas:

1. **Better HTML to WordML conversion**
2. **Support for tables and images**
3. **Multiple file download as ZIP**
4. **Dark mode toggle**
5. **More export formats (PDF, RTF)**

## 📄 License

MIT License - feel free to use this for any purpose!

## 🙏 Credits

- [Marked.js](https://marked.js.org/) - Markdown parser
- [PizZip](https://stuk.github.io/jszip/) - ZIP file creation
- [FileSaver.js](https://github.com/eligrey/FileSaver.js/) - File downloads
- Inspired by the Python version with pandoc

---

**Made with ❤️ for the markdown community**

*Host it once, use it forever! 🚀*
