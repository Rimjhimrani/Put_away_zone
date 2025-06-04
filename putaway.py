import streamlit as st
import pandas as pd
import os
from reportlab.lib.pagesizes import landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak, Image
from reportlab.lib.units import cm, inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.lib.utils import ImageReader
from io import BytesIO
import re
import tempfile
import subprocess
import sys

# Check for PIL and install if needed
try:
    from PIL import Image as PILImage
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    st.warning("Installing PIL...")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pillow'])
    from PIL import Image as PILImage
    PIL_AVAILABLE = True

# Check for QR code library and install if needed
try:
    import qrcode
    QR_AVAILABLE = True
except ImportError:
    QR_AVAILABLE = False
    st.warning("Installing QR code library...")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'qrcode'])
    import qrcode
    QR_AVAILABLE = True

# Define sticker dimensions
STICKER_WIDTH = 10 * cm
STICKER_HEIGHT = 15 * cm
STICKER_PAGESIZE = (STICKER_WIDTH, STICKER_HEIGHT)

# Define content box dimensions
CONTENT_BOX_WIDTH = 10 * cm  # Same width as page
CONTENT_BOX_HEIGHT = 7.2 * cm  # Half the page height

# Define paragraph styles
bold_style = ParagraphStyle(name='Bold', fontName='Helvetica-Bold', fontSize=16, alignment=TA_CENTER, leading=14)
desc_style = ParagraphStyle(name='Description', fontName='Helvetica', fontSize=11, alignment=TA_CENTER, leading=12)
qty_style = ParagraphStyle(name='Quantity', fontName='Helvetica', fontSize=11, alignment=TA_CENTER, leading=12)

def generate_qr_code(data_string):
    """Generate a QR code from the given data string"""
    try:
        # Create QR code instance
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=10,
            border=4,
        )
        
        # Add data
        qr.add_data(data_string)
        qr.make(fit=True)
        
        # Create QR code image
        qr_img = qr.make_image(fill_color="black", back_color="white")
        
        # Convert PIL image to bytes that reportlab can use
        img_buffer = BytesIO()
        qr_img.save(img_buffer, format='PNG')
        img_buffer.seek(0)
        
        # Create a QR code image with specified size
        return Image(img_buffer, width=2.5*cm, height=2.5*cm)
    except Exception as e:
        st.error(f"Error generating QR code: {e}")
        return None

def parse_location_string(location_str):
    """Parse a location string into components for table display"""
    # Initialize with empty values
    location_parts = [''] * 7

    if not location_str or not isinstance(location_str, str):
        return location_parts

    # Remove any extra spaces
    location_str = location_str.strip()

    # Try to parse location components
    pattern = r'([^_\s]+)'
    matches = re.findall(pattern, location_str)

    # Fill the available parts
    for i, match in enumerate(matches[:7]):
        location_parts[i] = match

    return location_parts

def generate_sticker_labels(df, date_width_ratio=0.5, date_height=1.5, qr_height=2.7):
    """Generate sticker labels with QR code from DataFrame"""
    
    # Create a function to draw the border box around content
    def draw_border(canvas, doc):
        canvas.saveState()
        # Draw border box around the content area (10cm x 7.5cm)
        # Position it at the top of the page with minimal margin
        x_offset = (STICKER_WIDTH - CONTENT_BOX_WIDTH) / 2
        y_offset = STICKER_HEIGHT - CONTENT_BOX_HEIGHT - 0.2*cm  # Position at top with minimal margin
        canvas.setStrokeColor(colors.Color(0, 0, 0, alpha=0.95))  # Slightly darker black (95% opacity)
        canvas.setLineWidth(1.8)  # Slightly thicker border
        canvas.rect(
            x_offset + doc.leftMargin,
            y_offset,
            CONTENT_BOX_WIDTH - 0.2*cm,  # Account for margins
            CONTENT_BOX_HEIGHT
        )
        canvas.restoreState()

    # Identify columns (case-insensitive)
    df.columns = [col.upper() if isinstance(col, str) else col for col in df.columns]
    cols = df.columns.tolist()

    # Find GRN No. column
    grn_col = next((col for col in cols if 'GRN' in col and ('NO' in col or 'NUM' in col or '#' in col)),
                   next((col for col in cols if col in ['GRN', 'GRNNO', 'GRN_NO']), 
                        next((col for col in cols if 'GOODS' in col and 'RECEIPT' in col), None)))

    # Find relevant columns
    part_no_col = next((col for col in cols if 'PART' in col and ('NO' in col or 'NUM' in col or '#' in col)),
                   next((col for col in cols if col in ['PARTNO', 'PART']), cols[0]))

    desc_col = next((col for col in cols if 'DESC' in col),
                   next((col for col in cols if 'NAME' in col), cols[1] if len(cols) > 1 else part_no_col))

    # Look for put away zone/location column
    put_away_col = next((col for col in cols if 'PUT' in col and 'AWAY' in col),
                       next((col for col in cols if 'PUTAWAY' in col),
                            next((col for col in cols if 'LOC' in col or 'POS' in col or 'LOCATION' in col),
                                 cols[2] if len(cols) > 2 else desc_col)))

    # Look for receipt date column
    receipt_date_col = next((col for col in cols if 'RECEIPT' in col and 'DATE' in col),
                           next((col for col in cols if 'RECEIPTDATE' in col or 'RECEIPT_DATE' in col),
                                next((col for col in cols if 'DATE' in col), None)))

    # Display column mapping info
    st.info(f"""
    **Column Mapping:**
    - GRN No: {grn_col if grn_col else 'Not found'}
    - Part No: {part_no_col}
    - Description: {desc_col}
    - Put Away Zone/Location: {put_away_col}
    - Receipt Date: {receipt_date_col if receipt_date_col else 'Not found'}
    """)

    # Create temporary file for PDF
    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_pdf.close()

    # Create document with minimal margins
    doc = SimpleDocTemplate(temp_pdf.name, pagesize=STICKER_PAGESIZE,
                          topMargin=0.2*cm,  # Minimal top margin
                          bottomMargin=(STICKER_HEIGHT - CONTENT_BOX_HEIGHT - 0.2*cm),  # Adjust bottom margin accordingly
                          leftMargin=0.1*cm, rightMargin=0.1*cm)

    content_width = CONTENT_BOX_WIDTH - 0.2*cm
    all_elements = []

    # Progress bar
    progress_bar = st.progress(0)
    status_text = st.empty()

    # Process each row as a single sticker
    total_rows = len(df)
    for index, row in df.iterrows():
        # Update progress
        progress = (index + 1) / total_rows
        progress_bar.progress(progress)
        status_text.text(f"Creating sticker {index+1} of {total_rows} ({int(progress*100)}%)")
        
        elements = []

        # Extract data
        grn_no = str(row[grn_col]) if grn_col and grn_col in row and pd.notna(row[grn_col]) else ""
        part_no = str(row[part_no_col])
        desc = str(row[desc_col])
        put_away_zone = str(row[put_away_col]) if put_away_col and put_away_col in row else ""
        receipt_date = str(row[receipt_date_col]) if receipt_date_col and receipt_date_col in row and pd.notna(row[receipt_date_col]) else ""
        
        location_parts = parse_location_string(put_away_zone)

        # Define row heights (increased slightly)
        grn_row_height = 0.9*cm      # New row for GRN No.
        header_row_height = 0.9*cm
        desc_row_height = 1.2*cm
        location_row_height = 1.0*cm

        # Clean receipt date - remove time if present
        clean_receipt_date = ""
        if receipt_date and receipt_date != "nan":
            # Try to extract just the date part
            try:
                if " " in receipt_date:
                    clean_receipt_date = receipt_date.split(" ")[0]
                else:
                    clean_receipt_date = receipt_date
            except:
                clean_receipt_date = receipt_date

        # Generate QR code with part information including GRN
        qr_data = f"GRN No: {grn_no}\nPart No: {part_no}\nDescription: {desc}\nPut Away Zone/Location: {put_away_zone}\nReceipt Date: {clean_receipt_date}"
        
        qr_image = generate_qr_code(qr_data)

        # Main table data (3 boxes: GRN No, Part No, Description)
        main_table_data = [
            ["GRN No", Paragraph(f"{grn_no}", bold_style)],
            ["Part No", Paragraph(f"{part_no}", bold_style)],
            ["Description", Paragraph(desc[:47] + "..." if len(desc) > 50 else desc, desc_style)]
        ]

        # Create main table
        main_table = Table(main_table_data,
                         colWidths=[content_width/3, content_width*2/3],
                         rowHeights=[grn_row_height, header_row_height, desc_row_height])

        main_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),  # Darker grid lines
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (0, -1), 11),
        ]))

        elements.append(main_table)

        # Put Away Zone/Location detail section
        put_away_label = Paragraph("Put Away Zone/Loc", ParagraphStyle(
            name='PutAwayLoc', fontName='Helvetica-Bold', fontSize=11, alignment=TA_CENTER
        ))

        # Total width for the 7 inner columns (2/3 of full content width)
        inner_table_width = content_width * 2 / 3
        
        # Define proportional widths - same as original 7 boxes
        col_proportions = [1.25, 1.25, 1.25, 1.25, 1, 1, 0.9]
        total_proportion = sum(col_proportions)
        
        # Calculate column widths based on proportions 
        inner_col_widths = [w * inner_table_width / total_proportion for w in col_proportions]

        put_away_inner_table = Table(
            [location_parts],
            colWidths=inner_col_widths,
            rowHeights=[location_row_height]
        )

        put_away_inner_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),  # Darker grid lines
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),  # Make location values bold
            ('FONTSIZE', (0, 0), (-1, -1), 9),
        ]))

        put_away_table = Table(
            [[put_away_label, put_away_inner_table]],
            colWidths=[content_width/3, inner_table_width],
            rowHeights=[location_row_height]
        )

        put_away_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),  # Darker grid lines
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))

        elements.append(put_away_table)

        # Bottom section - Receipt Date and QR Code in separate boxes in the same row
        # Use configurable dimensions from GUI
        date_row_height = date_height * cm      # Date section height (from GUI)
        qr_row_height = qr_height * cm          # QR code height (from GUI)
        
        # Date box width (left side) - configurable ratio from GUI
        date_width = content_width * date_width_ratio  # From GUI slider
        
        # QR code width (right side) - remaining width
        qr_width = content_width - date_width

        # Create Receipt Date box with label and date in horizontal layout
        date_table = Table(
            [["Receipt Date:", Paragraph(str(clean_receipt_date), qty_style)]],
            colWidths=[date_width*0.4, date_width*0.6],  # Split date box horizontally
            rowHeights=[date_row_height]
        )

        date_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN', (0, 0), (0, 0), 'RIGHT'),  # Right align "Receipt Date:" label
            ('ALIGN', (1, 0), (1, 0), 'LEFT'),   # Left align date value
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),  # Bold for label
            ('FONTSIZE', (0, 0), (0, 0), 10),
            ('FONTSIZE', (1, 0), (1, 0), 10),  # Font size for date value
        ]))

        # Create QR Code box - INVISIBLE BORDER
        if qr_image:
            qr_table = Table(
                [[qr_image]],
                colWidths=[qr_width],
                rowHeights=[qr_row_height]
            )
        else:
            qr_table = Table(
                [[Paragraph("QR", ParagraphStyle(
                    name='QRPlaceholder', fontName='Helvetica-Bold', fontSize=12, alignment=TA_CENTER
                ))]],
                colWidths=[qr_width],
                rowHeights=[qr_row_height]
            )

        # QR table style with INVISIBLE BORDER (using transparent/white color)
        qr_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0, colors.Color(1, 1, 1, alpha=0)),  # Invisible border (white with 0% opacity)
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))

        # Create the combined bottom row table with Date on left and QR on right
        # Use the taller height (QR height) for the combined table
        bottom_table = Table(
            [[date_table, qr_table]],
            colWidths=[date_width, qr_width],
            rowHeights=[qr_row_height]  # Use QR height as the overall row height
        )

        bottom_table.setStyle(TableStyle([
            # No grid for the outer table to avoid double borders
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Align to top since heights are different
        ]))

        # Add spacer before bottom section
        elements.append(Spacer(1, 0.3*cm))
        
        # Add the combined bottom table
        elements.append(bottom_table)

        # Add all elements for this sticker to the document
        all_elements.extend(elements)

        # Add page break after each sticker (except the last one)
        if index < len(df) - 1:
            all_elements.append(PageBreak())

    # Build the document
    try:
        # Pass the draw_border function to build to add border box
        doc.build(all_elements, onFirstPage=draw_border, onLaterPages=draw_border)
        progress_bar.progress(1.0)
        status_text.text("PDF generated successfully!")
        return temp_pdf.name
    except Exception as e:
        st.error(f"Error building PDF: {e}")
        return None

def main():
    st.set_page_config(
        page_title="Sticker Label Generator",
        page_icon="🏷️",
        layout="wide"
    )
    
    st.title("🏷️ Sticker Label Generator")
    st.markdown("Generate professional sticker labels with QR codes from your Excel/CSV data")
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("📐 Layout Configuration")
        
        # Date width configuration
        st.subheader("Date Width")
        date_width_percent = st.slider("Date Width (% of total)", 20, 80, 50, 5)
        date_width_ratio = date_width_percent / 100.0
        
        # Date height configuration
        st.subheader("Date Height")
        date_height = st.slider("Date Height (cm)", 1.0, 3.0, 1.5, 0.1)
        
        # QR height configuration
        st.subheader("QR Code Height")
        qr_height = st.slider("QR Height (cm)", 1.5, 4.0, 2.7, 0.1)
        
        st.markdown("---")
        st.markdown("**Preview Settings:**")
        st.text(f"Date Width: {date_width_percent}%")
        st.text(f"Date Height: {date_height} cm")
        st.text(f"QR Height: {qr_height} cm")
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📁 File Upload")
        uploaded_file = st.file_uploader(
            "Choose your Excel or CSV file",
            type=['xlsx', 'xls', 'csv'],
            help="Upload an Excel (.xlsx, .xls) or CSV file containing your sticker data"
        )
        
        if uploaded_file is not None:
            try:
                # Read the file
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)
                
                st.success(f"File loaded successfully! Found {len(df)} rows and {len(df.columns)} columns.")
                
                # Show data preview
                st.subheader("📊 Data Preview")
                st.dataframe(df.head(10), use_container_width=True)
                
                # Show column information
                st.subheader("📋 Column Information")
                col_info = pd.DataFrame({
                    'Column Name': df.columns,
                    'Data Type': df.dtypes,
                    'Non-Null Count': df.count(),
                    'Sample Value': [str(df[col].iloc[0]) if len(df) > 0 else '' for col in df.columns]
                })
                st.dataframe(col_info, use_container_width=True)
                
            except Exception as e:
                st.error(f"Error reading file: {str(e)}")
                df = None
        else:
            df = None
    
    with col2:
        st.header("⚙️ Generation")
        
        if df is not None:
            st.metric("Total Records", len(df))
            st.metric("Total Columns", len(df.columns))
            
            # Generate button
            if st.button("🚀 Generate Sticker Labels", type="primary", use_container_width=True):
                with st.spinner("Generating sticker labels..."):
                    pdf_path = generate_sticker_labels(
                        df, 
                        date_width_ratio=date_width_ratio,
                        date_height=date_height,
                        qr_height=qr_height
                    )
                    
                    if pdf_path:
                        # Read the PDF file
                        with open(pdf_path, 'rb') as pdf_file:
                            pdf_data = pdf_file.read()
                        
                        # Clean up temporary file
                        os.unlink(pdf_path)
                        
                        # Provide download button
                        st.success("✅ Sticker labels generated successfully!")
                        
                        filename = f"sticker_labels_{uploaded_file.name.split('.')[0]}.pdf"
                        st.download_button(
                            label="📥 Download PDF",
                            data=pdf_data,
                            file_name=filename,
                            mime="application/pdf",
                            use_container_width=True
                        )
                        
                        st.balloons()
                    else:
                        st.error("❌ Failed to generate sticker labels. Please check your data and try again.")
        else:
            st.info("👆 Please upload a file to begin")
    
    # Instructions
    with st.expander("📖 Instructions & Requirements"):
        st.markdown("""
        ### How to use this tool:
        1. **Upload your file**: Choose an Excel (.xlsx, .xls) or CSV file containing your sticker data
        2. **Configure layout**: Use the sidebar to adjust dimensions and layout
        3. **Generate labels**: Click the generate button to create your PDF
        4. **Download**: Download the generated PDF file
        
        ### Required/Expected Columns:
        The tool will automatically detect columns based on these patterns (case-insensitive):
        
        - **GRN No**: Columns containing "GRN" and ("NO" or "NUM" or "#")
        - **Part No**: Columns containing "PART" and ("NO" or "NUM" or "#")
        - **Description**: Columns containing "DESC" or "NAME"
        - **Put Away Zone/Location**: Columns containing "PUT AWAY", "PUTAWAY", "LOC", "POS", or "LOCATION"
        - **Receipt Date**: Columns containing "RECEIPT" and "DATE", or just "DATE"
        
        ### Features:
        - ✅ Automatic column detection
        - ✅ QR code generation with all item details
        - ✅ Customizable layout dimensions
        - ✅ Professional sticker format (10cm x 15cm)
        - ✅ Bordered content area for precise printing
        - ✅ Support for Excel and CSV files
        """)

if __name__ == "__main__":
    main()
