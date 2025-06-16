## 阿里巴巴1688订单转label打印

# Purchase to Printable PDF Converter

This script converts 1688 purchase order PDFs into printable labels with sequential numbering. It's designed to handle Chinese characters and create small format labels (4.9cm x 2.9cm) suitable for printing.

## Requirements

- Python 3.x
- Required Python packages (install using `pip install -r requirements.txt`):
  - pdfplumber>=0.7.0
  - reportlab>=3.6.0

## Installation

1. Clone or download this repository

2. Install the required packages:
```bash
pip install -r requirements.txt
```

3. Ensure you have a Chinese font installed on your system. The script will look for fonts in these locations:
   - macOS:
     - `/System/Library/Fonts/PingFang.ttc`
     - `/System/Library/Fonts/STHeiti Light.ttc`
     - `/System/Library/Fonts/STHeiti Medium.ttc`
     - `/System/Library/Fonts/Hiragino Sans GB.ttc`
   - Linux: `/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf`
   - Windows:
     - `C:\Windows\Fonts\simhei.ttf`
     - `C:\Windows\Fonts\msyh.ttc`

## Usage

1. Run the script:
```bash
python purchase_to_printable.py
```

2. When prompted, enter the path to your 1688 order PDF file.

3. Enter the starting number in the format `XXXX-YY` where:
   - XXXX is a 4-digit prefix (e.g., 2024 for year 2024)
   - YY is a 2-digit starting number (e.g., 01)

## Output

The script will generate a new PDF file with the same name as your input file, appended with "_output.pdf". Each page in the output PDF will contain:
- Product name (summarized to 10 words or less)
- Specifications (text after colon will be extracted)
- Quantity fraction (e.g., "1/5")
- Sequential number in the format XXXX-YY

## Features

- Automatically extracts tables from 1688 order PDF
- Handles Chinese characters
- Creates small format labels (4.9cm x 2.9cm)
- Automatically adjusts font size to fit content
- Generates sequential numbers for each label
- Supports multiple quantities per product
- Extracts text after colons in specifications
- Summarizes product names to fit label size

## Troubleshooting

If you encounter font-related errors:
1. Ensure you have a Chinese font installed on your system
2. Check if the font paths in the script match your system's font locations
3. Try installing additional Chinese fonts if needed

Common issues:
- Font not found: Install a Chinese font or verify font paths
- PDF table extraction fails: Ensure the PDF is not password protected
- Invalid number format: Use the correct XXXX-YY format

## Notes

- The script requires tables in the 1688 order PDF with columns for:
  - "货品名称" (Product Name)
  - "规格" (Specifications)
  - "数量" (Quantity)
- Text after colons in specifications will be extracted
- Product names will be summarized to 10 words or less
- Each product will generate multiple pages based on its quantity
- The output PDF is optimized for label printing (4.9cm x 2.9cm)
