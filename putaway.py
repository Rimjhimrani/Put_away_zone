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
import base64

# Auto-install required packages
try:
    from PIL import Image as PILImage
except ImportError:
    st.error("PIL not available. Please install: pip install pillow")
    st.stop()

try:
    import qrcode
except ImportError:
    st.error("qrcode not available. Please install: pip install qrcode")
    st.stop()

# Define sticker dimensions
STICKER_WIDTH = 10 * cm
STICKER_HEIGHT = 15 * cm
STICKER_PAGESIZE = (STICKER_WIDTH, STICKER_HEIGHT)

# Define content box dimensions
CONTENT_BOX_WIDTH = 10 * cm
CONTENT_BOX_HEIGHT = 7.2 * cm

# Define paragraph styles
bold_style = ParagraphStyle(name='Bold', fontName='Helvetica-Bold', fontSize=16, alignment=TA_CENTER, leading=14)
desc_style = ParagraphStyle(name='Description', fontName='Helvetica', fontSize=11, alignment=TA_CENTER, leading=12)
qty_style = ParagraphStyle(name='Quantity', fontName='Helvetica', fontSize=11, alignment=TA_CENTER, leading=12)

def generate_qr_code(data_string):
    """Generate a QR code from the given data string"""
    try:
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=10,
            border=4,
        )
        
        qr.add_data(data_string)
        qr.make(fit=True)
        
        qr_img = qr.make_image(fill_color="black", back_color="white")
        
        img_buffer = BytesIO()
        qr_img.save(img_buffer, format='PNG')
        img_buffer.seek(0)
        
        return Image(img_buffer, width=2.5*cm, height=2.5*cm)
    except Exception as e:
        st.error(f"Error generating QR code: {e}")
        return None

def parse_location_string(location_str):
    """Parse a location string into components for table display - only 4 boxes"""
    location_parts = [''] * 4
    if not location_str or not isinstance(location_str, str):
        return location_parts

    location_str = location_str.strip()
    pattern = r'([^_\s]+)'
    matches = re.findall(pattern, location_str)

    for i, match in enumerate(matches[:4]):
        location_parts[i] = match

    return location_parts

def generate_sticker_labels(df, date_width_ratio=0.5, date_height=1.5, qr_height=2.7):
    """Generate sticker labels with QR code from DataFrame"""
    
    def draw_border(canvas, doc):
        canvas.saveState()
        x_offset = (STICKER_WIDTH - CONTENT_BOX_WIDTH) / 2
        y_offset = STICKER_HEIGHT - CONTENT_BOX_HEIGHT - 0.2*cm
        canvas.setStrokeColor(colors.Color(0, 0, 0, alpha=0.95))
        canvas.setLineWidth(1.8)
        canvas.rect(
            x_offset + doc.leftMargin,
            y_offset,
            CONTENT_BOX_WIDTH - 0.2*cm,
            CONTENT_BOX_HEIGHT
        )
        canvas.restoreState()

    # Identify columns (case-insensitive)
    original_columns = df.columns.tolist()
    df_copy = df.copy()
    df_copy.columns = [col.upper() if isinstance(col, str) else col for col in df_copy.columns]
    cols = df_copy.columns.tolist()

    # Find GRN No. column
    grn_col = next((col for col in cols if 'GRN' in col and ('NO' in col or 'NUM' in col or '#' in col)),
                   next((col for col in cols if col in ['GRN', 'GRNNO', 'GRN_NO']), 
                        next((col for col in cols if 'GOODS' in col and 'RECEIPT' in col), None)))

    # Find relevant columns
    part_no_col = next((col for col in cols if 'PART' in col and ('NO' in col or 'NUM' in col or '#' in col)),
                   next((col for col in cols if col in ['PARTNO', 'PART']), cols[0]))

    desc_col = next((col for col in cols if 'DESC' in col),
                   next((col for col in cols if 'NAME' in col), cols[1] if len(cols) > 1 else part_no_col))

    # Look for store location column
    store_location_col = next((col for col in cols if 'STORE' in col and 'LOC' in col),
                             next((col for col in cols if 'STORELOCATION' in col or 'STORE_LOCATION' in col),
                                  next((col for col in cols if 'LOC' in col or 'POS' in col or 'LOCATION' in col),
                                       cols[2] if len(cols) > 2 else desc_col)))

    # Look for receipt date column
    receipt_date_col = next((col for col in cols if 'RECEIPT' in col and 'DATE' in col),
                           next((col for col in cols if 'RECEIPTDATE' in col or 'RECEIPT_DATE' in col),
                                next((col for col in cols if 'DATE' in col), None)))

    # Create temporary PDF file
    temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_pdf_path = temp_pdf.name
    temp_pdf.close()

    # Create document with minimal margins
    doc = SimpleDocTemplate(temp_pdf_path, pagesize=STICKER_PAGESIZE,
                          topMargin=0.2*cm,
                          bottomMargin=(STICKER_HEIGHT - CONTENT_BOX_HEIGHT - 0.2*cm),
                          leftMargin=0.1*cm, rightMargin=0.1*cm)

    content_width = CONTENT_BOX_WIDTH - 0.2*cm
    all_elements = []

    # Progress bar
    progress_bar = st.progress(0)
    status_placeholder = st.empty()
    
    # Process each row as a single sticker
    total_rows = len(df_copy)
    for index, row in df_copy.iterrows():
        # Update progress
        progress = (index + 1) / total_rows
        progress_bar.progress(progress)
        status_placeholder.text(f"Creating sticker {index+1} of {total_rows} ({int(progress*100)}%)")
        
        elements = []

        # Extract data
        grn_no = str(row[grn_col]) if grn_col and grn_col in row and pd.notna(row[grn_col]) else ""
        part_no = str(row[part_no_col])
        desc = str(row[desc_col])
        store_location = str(row[store_location_col]) if store_location_col and store_location_col in row else ""
        receipt_date = str(row[receipt_date_col]) if receipt_date_col and receipt_date_col in row and pd.notna(row[receipt_date_col]) else ""
        
        location_parts = parse_location_string(store_location)

        # Define row heights
        grn_row_height = 0.9*cm
        header_row_height = 0.9*cm
        desc_row_height = 1.4*cm
        location_row_height = 0.8*cm

        # Clean receipt date
        clean_receipt_date = ""
        if receipt_date and receipt_date != "nan":
            try:
                if " " in receipt_date:
                    clean_receipt_date = receipt_date.split(" ")[0]
                else:
                    clean_receipt_date = receipt_date
            except:
                clean_receipt_date = receipt_date

        # Generate QR code
        qr_data = f"GRN No: {grn_no}\nPart No: {part_no}\nDescription: {desc}\nStore Location: {store_location}\nReceipt Date: {clean_receipt_date}"
        qr_image = generate_qr_code(qr_data)

        # Main table data
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
            ('GRID', (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (0, -1), 11),
        ]))

        elements.append(main_table)

        # Store Location section
        store_location_label = Paragraph("Store Location", ParagraphStyle(
            name='StoreLocation', fontName='Helvetica-Bold', fontSize=11, alignment=TA_CENTER
        ))

        inner_table_width = content_width * 2 / 3
        inner_col_widths = [inner_table_width / 4] * 4

        store_location_inner_table = Table(
            [location_parts],
            colWidths=inner_col_widths,
            rowHeights=[location_row_height]
        )

        store_location_inner_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
        ]))

        store_location_table = Table(
            [[store_location_label, store_location_inner_table]],
            colWidths=[content_width/3, inner_table_width],
            rowHeights=[location_row_height]
        )

        store_location_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))

        elements.append(store_location_table)

        # Bottom section
        date_row_height = date_height * cm
        qr_row_height = qr_height * cm
        date_width = content_width * date_width_ratio
        qr_width = content_width - date_width

        # Create Receipt Date box
        date_table = Table(
            [["Receipt Date:", Paragraph(str(clean_receipt_date), qty_style)]],
            colWidths=[date_width*0.4, date_width*0.6],
            rowHeights=[date_row_height]
        )

        date_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1.2, colors.Color(0, 0, 0, alpha=0.95)),
            ('ALIGN', (0, 0), (0, 0), 'RIGHT'),
            ('ALIGN', (1, 0), (1, 0), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (0, 0), 10),
            ('FONTSIZE', (1, 0), (1, 0), 10),
        ]))

        # Create QR Code box
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

        qr_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0, colors.Color(1, 1, 1, alpha=0)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))

        # Combined bottom table
        bottom_table = Table(
            [[date_table, qr_table]],
            colWidths=[date_width, qr_width],
            rowHeights=[qr_row_height]
        )

        bottom_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))

        elements.append(Spacer(1, 0.3*cm))
        elements.append(bottom_table)

        all_elements.extend(elements)

        # Add page break except for last sticker
        if index < len(df_copy) - 1:
            all_elements.append(PageBreak())

    # Build the document
    try:
        doc.build(all_elements, onFirstPage=draw_border, onLaterPages=draw_border)
        status_placeholder.text("PDF generated successfully!")
        progress_bar.progress(1.0)
        return temp_pdf_path
    except Exception as e:
        st.error(f"Error building PDF: {e}")
        return None

def get_download_link(file_path, filename):
    """Generate a download link for the PDF file"""
    with open(file_path, "rb") as f:
        bytes_data = f.read()
    b64 = base64.b64encode(bytes_data).decode()
    href = f'<a href="data:application/pdf;base64,{b64}" download="{filename}">Download PDF</a>'
    return href

def main():
    st.set_page_config(
        page_title="Sticker Label Generator",
        page_icon="🏷️",
        layout="wide"
    )
    
    st.title("🏷️ Sticker Label Generator")
    st.markdown("Generate professional sticker labels with QR codes from your Excel/CSV data")
    
    # Sidebar for settings
    st.sidebar.header("Layout Settings")
    
    # Date width control
    date_width_percent = st.sidebar.slider(
        "Date Width (%)", 
        min_value=20, 
        max_value=80, 
        value=50, 
        step=5,
        help="Percentage of total width for the date section"
    )
    date_width_ratio = date_width_percent / 100.0
    
    # Date height control
    date_height = st.sidebar.slider(
        "Date Height (cm)", 
        min_value=1.0, 
        max_value=3.0, 
        value=1.2, 
        step=0.1,
        help="Height of the date section"
    )
    
    # QR height control
    qr_height = st.sidebar.slider(
        "QR Height (cm)", 
        min_value=1.5, 
        max_value=4.0, 
        value=2.3, 
        step=0.1,
        help="Height of the QR code section"
    )
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("Upload File")
        uploaded_file = st.file_uploader(
            "Choose an Excel or CSV file",
            type=['xlsx', 'xls', 'csv'],
            help="Upload your data file containing the information for sticker labels"
        )
        
        if uploaded_file is not None:
            try:
                # Read the file
                if uploaded_file.name.lower().endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)
                
                st.success(f"✅ File loaded successfully! Found {len(df)} rows and {len(df.columns)} columns.")
                
                # Show column information
                st.subheader("📊 Data Preview")
                st.write(f"**Columns found:** {', '.join(df.columns.tolist())}")
                
                # Show first few rows
                st.write(df.head())
                
                # Generate button
                st.subheader("🎯 Generate Labels")
                
                if st.button("🚀 Generate Sticker Labels", type="primary"):
                    with st.spinner("Generating sticker labels..."):
                        pdf_path = generate_sticker_labels(
                            df, 
                            date_width_ratio=date_width_ratio,
                            date_height=date_height,
                            qr_height=qr_height
                        )
                        
                        if pdf_path:
                            st.success("🎉 Sticker labels generated successfully!")
                            
                            # Create download link
                            filename = f"{uploaded_file.name.split('.')[0]}_sticker_labels.pdf"
                            download_link = get_download_link(pdf_path, filename)
                            st.markdown(download_link, unsafe_allow_html=True)
                            
                            # Clean up temporary file after a delay
                            import time
                            time.sleep(1)
                            try:
                                os.unlink(pdf_path)
                            except:
                                pass
                        else:
                            st.error("❌ Failed to generate sticker labels.")
                            
            except Exception as e:
                st.error(f"❌ Error reading file: {str(e)}")
    
    with col2:
        st.header("ℹ️ Instructions")
        st.markdown("""
        **How to use:**
        
        1. **Upload your file** (Excel or CSV)
        2. **Adjust layout settings** in the sidebar
        3. **Click Generate** to create your sticker labels
        4. **Download** the generated PDF
        
        **Expected columns:**
        - GRN No./GRN Number
        - Part No./Part Number  
        - Description/Name
        - Store Location
        - Receipt Date
        
        **Features:**
        - ✅ QR codes with all item data
        - ✅ Professional border layout
        - ✅ Customizable dimensions
        - ✅ Automatic column detection
        - ✅ Ready for printing
        """)
        
        st.header("📋 Layout Preview")
        st.markdown(f"""
        **Current Settings:**
        - Date Width: {date_width_percent}%
        - Date Height: {date_height} cm
        - QR Height: {qr_height} cm
        - Sticker Size: 10×15 cm
        """)

if __name__ == "__main__":
    main()
