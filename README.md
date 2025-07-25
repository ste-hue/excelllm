# Excel to LLM Converter

A clean, minimal web application that converts Excel files into AI-friendly formats with formula extraction capabilities.

## Features

- 📊 **Excel File Processing**: Support for .xlsx and .xls files
- 📑 **Multi-Sheet Support**: Select and process individual sheets
- 🎯 **Range Selection**: Extract specific cell ranges (e.g., A1:Z100)
- 🧮 **Formula Extraction**: Automatically detects and extracts Excel formulas
- 🔄 **Multiple Output Formats**:
  - JSON - Structured data for APIs
  - CSV - Comma-separated values
  - Markdown - Tables for documentation
  - Text - Tab-delimited plain text
- 💡 **LLM-Optimized**: Includes metadata and context for better AI understanding
- 🚀 **Client-Side Processing**: No server needed, all processing happens in the browser

## Getting Started

### Prerequisites

- Node.js 16+ and npm

### Installation

```bash
# Clone the repository
git clone <your-repo-url>
cd excelllm

# Install dependencies
npm install

# Start development server
npm run dev
```

### Building for Production

```bash
npm run build
```

The production-ready files will be in the `dist` directory.

## Deployment

### Vercel

The project includes a `vercel.json` configuration file. Simply:

1. Install Vercel CLI: `npm i -g vercel`
2. Run: `vercel`
3. Follow the prompts

### Netlify

The project includes a `netlify.toml` configuration file. You can:

1. Push to a Git repository
2. Connect the repository to Netlify
3. Deploy automatically

Or use the Netlify CLI:

```bash
npm i -g netlify-cli
netlify deploy --prod
```

### Other Static Hosts

Since this is a client-side only application, you can deploy the `dist` folder to any static hosting service:

- GitHub Pages
- Cloudflare Pages
- AWS S3 + CloudFront
- Firebase Hosting

## Usage

1. **Upload File**: Drag and drop or click to upload an Excel file (max 10MB)
2. **Select Sheet**: Choose which sheet to process from the dropdown
3. **Set Range**: Specify the cell range to extract (default: A1:Z100)
4. **Process**: Click "Process Sheet" to convert the data
5. **Export**: Choose your desired format and download

## Technical Details

### Built With

- **React** - UI framework
- **TypeScript** - Type safety
- **Vite** - Build tool
- **Tailwind CSS** - Styling
- **XLSX** - Excel file processing

### Performance Optimizations

- Code splitting for XLSX library
- Lazy loading of large dependencies
- Minimal bundle size
- Client-side only (no server overhead)

## Security Considerations

- All processing happens client-side
- No data is sent to any server
- Files are processed in memory only
- No data persistence or tracking

## Browser Support

- Chrome 90+
- Firefox 88+
- Safari 14+
- Edge 90+

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is MIT licensed. See LICENSE file for details.

## Acknowledgments

- [SheetJS](https://sheetjs.com/) for the excellent XLSX library
- [Tailwind CSS](https://tailwindcss.com/) for the utility-first CSS framework
- [Lucide React](https://lucide.dev/) for the icon set