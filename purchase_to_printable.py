import pdfplumber
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, PageBreak
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth
import os

def extract_after_colon(text):
    """Extract text after the colon for each line, or return original text if no colon found."""
    if not text:
        return ""

    # Split by newlines and process each line
    lines = text.split('\n')
    processed_lines = []

    for line in lines:
        line = line.strip()
        if '：' in line:
            # Get text after the last colon in the line
            parts = line.split('：')
            processed_lines.append(parts[-1].strip())
        else:
            processed_lines.append(line)

    # Join the processed lines with newlines
    return '\n'.join(processed_lines)

def summarize_text(text, max_words=10):
    """Summarize text to have less than max_words."""
    if not text:
        return ""

    # Split into words and remove empty strings
    words = [word.strip() for word in text.split() if word.strip()]

    # If already less than max_words, return as is
    if len(words) <= max_words:
        return text

    # Take first max_words words
    return ' '.join(words[:max_words])

def calculate_font_size_for_columns(name, spec, quantity, max_width, font_name, start_size=16):
    """Calculate font size to fit all three columns on one page."""
    current_size = start_size

    while current_size > 8:  # Don't go smaller than 8pt
        try:
            # Calculate width for each column
            name_width = stringWidth(name, font_name, current_size)
            spec_width = stringWidth(spec, font_name, current_size)
            quantity_width = stringWidth(quantity, font_name, current_size)

            # Check if all columns fit within max_width
            if name_width <= max_width and spec_width <= max_width and quantity_width <= max_width:
                break
        except:
            # If there's an error with font metrics, try a smaller size
            pass

        current_size -= 1

    return current_size

# Register Chinese font
def register_chinese_font():
    # Try to find a Chinese font in common locations
    font_paths = [
        '/System/Library/Fonts/PingFang.ttc',  # macOS
        '/System/Library/Fonts/STHeiti Light.ttc',  # macOS
        '/System/Library/Fonts/STHeiti Medium.ttc',  # macOS
        '/System/Library/Fonts/Hiragino Sans GB.ttc',  # macOS
        '/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf',  # Linux
        'C:\\Windows\\Fonts\\simhei.ttf',  # Windows
        'C:\\Windows\\Fonts\\msyh.ttc',  # Windows
    ]

    for font_path in font_paths:
        if os.path.exists(font_path):
            try:
                # Try to register the font
                pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                # Test the font with a simple string
                stringWidth('测试', 'ChineseFont', 12)
                return True
            except Exception as e:
                print(f"Warning: Could not use font {font_path}: {str(e)}")
                continue

    # If no system font is found, use a default font
    print("Warning: No suitable Chinese font found. Using default font which may not support Chinese characters.")
    return False

def find_target_tables(tables):
    """Find all tables with the required columns from a list of tables."""
    target_tables = []
    headers = None

    for table in tables:
        if table and len(table) > 0:
            current_headers = [str(cell).strip() if cell else '' for cell in table[0]]
            if '货品名称' in current_headers and '规格' in current_headers and '数量' in current_headers:
                target_tables.append(table)
                if headers is None:
                    headers = current_headers

    return target_tables, headers

def get_user_input():
    """Get the starting number from user input."""
    while True:
        try:
            user_input = input("Please enter the starting number (format: XXXX-YY): ")
            if '-' not in user_input:
                print("Please use the format XXXX-YY")
                continue

            prefix, start_num = user_input.split('-')
            if len(prefix) != 4 or not prefix.isdigit():
                print("Prefix must be 4 digits")
                continue

            if not start_num.isdigit() or len(start_num) != 2:
                print("Start number must be 2 digits")
                continue

            return prefix, int(start_num)
        except ValueError:
            print("Invalid input. Please use the format XXXX-YY")

def create_pdf_from_pdf(input_pdf, output_pdf):
    # Get user input for starting number
    prefix, start_num = get_user_input()

    # Register Chinese font
    has_chinese_font = register_chinese_font()

    # Read the PDF file
    try:
        with pdfplumber.open(input_pdf) as pdf:
            # Extract tables from all pages
            all_tables = []
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    all_tables.extend(tables)

            if not all_tables:
                print("No tables found in the PDF")
                return

            # Find all tables with the required columns
            target_tables, headers = find_target_tables(all_tables)
            if not target_tables:
                print("Required columns not found in the PDF")
                return

            # Get column indices
            name_idx = headers.index('货品名称')
            spec_idx = headers.index('规格')
            quantity_idx = headers.index('数量')

            # Define custom page size (4.9cm x 2.9cm)
            custom_page_size = (4.9 * cm, 2.9 * cm)

            # Create PDF document with custom page size and minimal margins
            doc = SimpleDocTemplate(
                output_pdf,
                pagesize=custom_page_size,
                rightMargin=0.1 * cm,
                leftMargin=0.1 * cm,
                topMargin=0.1 * cm,
                bottomMargin=0.1 * cm
            )

            # Get base styles
            styles = getSampleStyleSheet()

            # Set font name
            font_name = 'ChineseFont' if has_chinese_font else 'Helvetica'

            # Calculate available width for text
            available_width = custom_page_size[0] - (0.2 * cm)  # Subtract minimal margins

            # Create content for each page
            content = []
            current_num = start_num

            # Process each table
            for target_table in target_tables:
                # Process each row (skip header row)
                for row in target_table[1:]:
                    if not row or len(row) <= max(name_idx, spec_idx, quantity_idx):
                        continue

                    name = str(row[name_idx]).strip() if row[name_idx] else ''
                    spec = str(row[spec_idx]).strip() if row[spec_idx] else ''
                    quantity = str(row[quantity_idx]).strip() if row[quantity_idx] else ''

                    if not name or not quantity:
                        continue

                    # Summarize the name to less than 10 words
                    name = summarize_text(name)

                    # Extract text after colon in spec
                    spec = extract_after_colon(spec)

                    try:
                        # Calculate appropriate font size for all columns
                        font_size = calculate_font_size_for_columns(name, spec, quantity, available_width, font_name)

                        # Create styles with calculated font size and minimal spacing
                        heading_style = ParagraphStyle(
                            'CustomHeading',
                            parent=styles['Heading1'],
                            fontName=font_name,
                            fontSize=font_size,
                            alignment=TA_CENTER,
                            spaceAfter=2,  # Reduced spacing
                            leading=font_size  # Set leading equal to font size
                        )

                        normal_style = ParagraphStyle(
                            'CustomNormal',
                            parent=styles['Normal'],
                            fontName=font_name,
                            fontSize=font_size,
                            alignment=TA_CENTER,
                            spaceAfter=2,  # Reduced spacing
                            leading=font_size  # Set leading equal to font size
                        )

                        bold_style = ParagraphStyle(
                            'BoldStyle',
                            parent=styles['Normal'],
                            fontName=font_name,
                            fontSize=font_size * 1.2,  # Make the number larger
                            alignment=TA_CENTER,
                            spaceAfter=2,
                            leading=font_size * 1.2,  # Adjust leading to match font size
                            textColor='black',
                            fontWeight='bold'
                        )

                        # Convert quantity to integer
                        try:
                            quantity_num = int(quantity)
                        except ValueError:
                            quantity_num = 1

                        # Create pages based on quantity
                        for i in range(quantity_num):
                            # Add all three columns with minimal spacing
                            content.append(Paragraph(name, heading_style))
                            content.append(Paragraph(spec, normal_style))

                            # Always show fraction format
                            fraction_text = f"{i + 1}/{quantity_num}"
                            content.append(Paragraph(fraction_text, normal_style))

                            # Add the increasing number
                            number_text = f"{prefix}-{current_num:02d}"
                            content.append(Paragraph(number_text, bold_style))
                            current_num += 1

                            # Add page break after each page
                            content.append(PageBreak())

                    except Exception as e:
                        print(f"Warning: Error processing row: {str(e)}")
                        continue

            # Build PDF
            doc.build(content)
            print(f"PDF has been created successfully: {output_pdf}")

    except Exception as e:
        print(f"Error processing PDF file: {e}")
        return

def main():
    # Get input file from user
    input_pdf = input("Please enter the path to your PDF file: ")

    # Validate file exists
    if not os.path.exists(input_pdf):
        print("File does not exist!")
        return

    # Create output filename
    output_pdf = os.path.splitext(input_pdf)[0] + "_output.pdf"

    # Process the file
    create_pdf_from_pdf(input_pdf, output_pdf)

if __name__ == "__main__":
    main()
