# DOCX Document Generator

A professional document generation script that creates formatted Word documents (.docx) with images, tables, headers, footers, numbered/bulleted lists, and styled text using Node.js and the `docx` library.

## Features

- ✅ ESM (ES Modules) support
- ✅ Professional document styling with custom fonts and colors
- ✅ Headers and footers with page numbers
- ✅ Embedded images with automatic scaling
- ✅ Tables with alternating row colors
- ✅ Numbered and bulleted lists
- ✅ Info boxes with borders and shading
- ✅ Multiple heading levels (H1, H2)
- ✅ Page breaks and dividers
- ✅ Custom spacing and indentation

## Requirements

- **Node.js** >= 18.0.0 or **Bun** >= 1.0.0
- **docx** library (included in dependencies)

## Installation

### Using npm
```bash
npm install
```

### Using Bun
```bash
bun install
```

## Usage

### Run with Bun (Recommended)
```bash
bun new.js
```

### Run with Node.js
```bash
node new.js
```

### Using npm script
```bash
npm start
```

## Output

The script generates a file named:
```
Assignment_2D_BSCS22054.docx
```

This file is created in the current working directory.

## Project Structure

```
.
├── new.js                              # Main document generation script
├── tmp/                                # Folder containing image assets (excluded from git)
│   ├── 2d_img1.png
│   ├── 2d_img2.png
│   ├── 2d_img3.png
│   └── 2d_img4.png
├── package.json                        # Project dependencies and scripts
├── README.md                           # This file
├── LICENSE                             # MIT License
└── .gitignore                          # Git ignore configuration
```

## Configuration

### Image Paths
Images are loaded from the `tmp/` folder relative to the script location:
- `tmp/2d_img1.png` - Board screenshot
- `tmp/2d_img2.png` - List view
- `tmp/2d_img3.png` - Velocity chart
- `tmp/2d_img4.png` - Burndown chart

### Document Settings
- **Page Size**: Letter (8.5" x 11")
- **Margins**: 1.2" on all sides
- **Default Font**: Arial, 11pt
- **Color Scheme**: Blue tones (#1E3A5F, #2E75B6)

## Dependencies

- **docx** (^8.12.0) - Create .docx files with Node.js

## License

MIT License - See [LICENSE](./LICENSE) file for details.

## Author

Syed Muhammad Abdullah (BSCS22054)

---

**Note**: The `tmp/` folder is excluded from Git. To use this script, ensure image files are present in the `tmp/` directory before running.
