import os
import io
import sys
import json
import qrcode
import PyPDF2
import tempfile
import datetime 
import pandas as pd 
from copy import copy 
from pathlib import Path 
from reportlab.pdfgen import canvas 
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.ttfonts import TTFont   
from reportlab.lib.pagesizes import letter 
  
# montserrat  
Montserrat_font_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'fonts', 'Montserrat.ttf')  
pdfmetrics.registerFont(TTFont("Montserrat-Regular", Montserrat_font_path))  

def match_sku_to_item_description(sku):    
    if getattr(sys, 'frozen', False):      
        bundle_dir = sys._MEIPASS      
    else:      
        bundle_dir = os.path.dirname(os.path.abspath(__file__))      
    
    json_path = Path(bundle_dir) / "sku_to_item_description.json"      
    
    # Load SKU to item description mapping from JSON file        
    with json_path.open() as f:        
        sku_to_item_description = json.load(f)     
    
    for key in sku_to_item_description:      
        if sku.startswith(key):      
            return sku_to_item_description[key]      
    return "Description not found"  

def count_skus(order_data):  
    sku_counts = {}  
    for row in order_data[1:]:  
        sku = row[2]  
        if sku not in sku_counts:  
            sku_counts[sku] = 0  
        sku_counts[sku] += 1  
    return sku_counts 

def wrap_text(text, max_width, canvas, font, font_size):  
    text_lines = []  
    current_line = []  
    words = text.split(" ")  
    for word in words:  
        current_line.append(word)  
        line_width = canvas.stringWidth(" ".join(current_line), font, font_size)  
        if line_width > max_width:  
            current_line.pop()  # Remove the last word  
            text_lines.append(" ".join(current_line))  
            current_line = [word]  # Start a new line with the last word  
    text_lines.append(" ".join(current_line))  
    return text_lines

def draw_sku_grid(c, sku_counts, sku_name_grid_coordinates, font_size_sku):    
    print("Before sorting:")      
    print(sku_counts)    
    
    # Set font and font color for the SKUs      
    c.setFont("Montserrat-Regular", font_size_sku)      
    c.setFillColorRGB(0, 0, 0) 

    def draw_sku_qr_code(c, sku, x, y, width, height):  
        qr = qrcode.QRCode(  
            version=1,  
            error_correction=qrcode.constants.ERROR_CORRECT_L,  
            box_size=2,  
            border=0,  
        )  
        qr.add_data(sku)  
        qr.make(fit=True)  
  
        qr_img = qr.make_image(fill_color="black", back_color="white")  
  
        temp_qr_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')  
        qr_img.save(temp_qr_file.name)  
  
        c.drawImage(ImageReader(temp_qr_file.name), x, y, width, height)  
  
        temp_qr_file.close()  
        os.remove(temp_qr_file.name)   
    
    # Custom sorting key for SKUs    
    def custom_sku_sort_key(sku_count):      
        sku, _ = sku_count      
        return sku    
    
    # Sort the SKUs alphabetically      
    sorted_skus = sorted(sku_counts.items(), key=custom_sku_sort_key)     
    
    print("After sorting:")    
    print(sorted_skus) 

    qr_code_grid_coordinates = [  
    (248, 636), (502, 636),  
    (248, 546), (502, 546),  
    (248, 456), (502, 456),  
    (248, 366), (502, 366),  
    (248, 276), (502, 276),  
    (248, 186), (502, 186),  
    (248, 96), (502, 96),  
    ]
    
    # Draw the SKUs and their counts on the canvas using grid coordinates      
    remaining_skus = len(sorted_skus)      
    for idx, (sku, count) in enumerate(sorted_skus):      
        if idx >= len(sku_name_grid_coordinates):      
            print(f"SKUs not drawn: {sorted_skus[idx:]}")    
            return remaining_skus - idx      
        x, y = sku_name_grid_coordinates[idx]      
        c.drawString(x, y, f"{count}x {sku}")    
        print(f"Drawing SKU: {sku}")

        # Draw the QR code for the current SKU next to its description using the QR code grid  
        qr_code_x, qr_code_y = qr_code_grid_coordinates[idx]  
        qr_code_width = 40  # Adjust the width of the QR code  
        qr_code_height = 40  # Adjust the height of the QR code  
        draw_sku_qr_code(c, sku, qr_code_x, qr_code_y, qr_code_width, qr_code_height)
  
        sku_description = match_sku_to_item_description(sku)    
          
        # Set the font size for the item description  
        font_size_description = 9 
          
        # Wrap the item description text with the new font size  
        wrapped_sku_description = wrap_text(str(sku_description), 175, c, "Montserrat-Regular", font_size_description)  # max_description_width = 200    
    
        x_desc, y_desc = sku_name_grid_coordinates[idx]    
        y_desc -= 15    
    
        # Draw the item description with the new font size  
        c.setFont("Montserrat-Regular", font_size_description)  
        for line in wrapped_sku_description:    
            c.drawString(x_desc, y_desc, line)    
            y_desc -= 15  
          
        # Reset the font size back to font_size_sku for the next iteration  
        c.setFont("Montserrat-Regular", font_size_sku)  
      
    return 0  
 
def create_pick_list_pdf(order_data, background_pdf_path, match_sku_to_item_description):      
    packet = io.BytesIO()      
    c = canvas.Canvas(packet, pagesize=letter)      
    temp_pdf = 'temp_pick_list.pdf'      
    custom_field_3 = order_data[1][5]
    final_pdf = os.path.join(output_folder_path, f'{custom_field_3}_pick_list.pdf')
    current_date = datetime.datetime.now().strftime('%m.%d.%Y %H:%M')


    # Get the current date and format it 
    def draw_date_and_custom_field_3(c, date, custom_field_3): 
             
        # Set font, font size, and font color for the date      
        font_size_date = 15      
        c.setFont("Montserrat-Regular", font_size_date)      
        c.setFillColorRGB(0, 0, 0)      
        # Draw the date      
        c.drawString(215, 761, current_date)      

        # Draw the custom field 3     
        font_size_custom_field = 15      
        c.setFont("Montserrat-Regular", font_size_custom_field)      
        c.setFillColorRGB(0, 0, 0)      
        c.drawString(215, 740, str(custom_field_3))

    # Draw the date and custom field 3 for the first page  
    draw_date_and_custom_field_3(c, current_date, custom_field_3)
  
    # Draw QR code function  
    def draw_qr_code(c, custom_field_3, x, y, width, height):  
        qr = qrcode.QRCode(  
            version=1,  
            error_correction=qrcode.constants.ERROR_CORRECT_L,  
            box_size=2,  
            border=0,  
        )  
        qr.add_data(str(custom_field_3))  # Use custom_field_3 value for the QR code  
        qr.make(fit=True)  
  
        # Create QR code image  
        qr_img = qr.make_image(fill_color="black", back_color="white")  
  
        # Save the QR code image to a temporary file  
        temp_qr_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')  
        qr_img.save(temp_qr_file.name)  
  
        # Draw the resized image on the canvas at the specified coordinates  
        c.drawImage(ImageReader(temp_qr_file.name), x, y, width, height)  
  
        # Close and remove the temporary QR code file  
        temp_qr_file.close()  
        os.remove(temp_qr_file.name) 

    # Draw the QR code for the first page  
    draw_qr_code(c, custom_field_3, 483, 730, 60, 60)  # (x, y, width, height) 
      
    # Count the SKUs  
    sku_counts = count_skus(order_data)  
  
    font_size_sku = 12
    c.setFont("Montserrat-Regular", font_size_sku)  
    c.setFillColorRGB(0, 0, 0)  
  
    sku_name_grid_coordinates = [    
        (65, 685), (310, 685),    
        (65, 595), (310, 595),    
        (65, 505), (310, 505),    
        (65, 415), (310, 415),    
        (65, 325), (310, 325),    
        (65, 235), (310, 235),    
        (65, 145), (310, 145),    
    ]    
    
    remaining_skus = draw_sku_grid(c, sku_counts, sku_name_grid_coordinates, font_size_sku)   
  
    # Calculate the total number of pages  
    total_pages = 1 + (len(sku_counts) - 1) // len(sku_name_grid_coordinates)  
  
    # Add page number function  
    def draw_page_number(c, page_number, total_pages, x, y):  
        c.setFont("Montserrat-Regular", 10)  
        c.setFillColorRGB(0, 0, 0)  
        c.drawString(x, y, f"Page {page_number} of {total_pages}") 

    # Draw the first page number  
    page_number = 1  
    draw_page_number(c, page_number, total_pages, 500, 50) 

    # Add more pages if there are more SKUs than slots available    
    while remaining_skus > 0:      
        c.showPage()     
        page_number += 1     
        draw_page_number(c, page_number, total_pages, 500, 50)     
        c.setFont("Montserrat-Regular", font_size_sku)        
        c.setFillColorRGB(0, 0, 0)      
        
        # Draw the date and custom field 3 for the new page      
        draw_date_and_custom_field_3(c, current_date, custom_field_3)    
        draw_qr_code(c, custom_field_3, 483, 730, 60, 60)  # (x, y, width, height)     
        
        # Update the sku_counts with the remaining SKUs  
        sku_counts = dict(list(sku_counts.items())[len(sku_name_grid_coordinates):])  
        
        # Draw the remaining SKUs on the new page 
        remaining_skus = draw_sku_grid(c, sku_counts, sku_name_grid_coordinates, font_size_sku)    
    
    # Save the canvas to 'temp_pick_list.pdf'  
    c.save()      
      
    packet.seek(0)      
    with open(temp_pdf, 'wb') as temp_file:      
        temp_file.write(packet.getbuffer())      
      
    # Merge the background PDF with the temp_pick_list.pdf  
    with open(background_pdf_path, 'rb') as background_file, open(temp_pdf, 'rb') as content_file:  
        background_pdf = PyPDF2.PdfReader(background_file)  
        content_pdf = PyPDF2.PdfReader(content_file)  
        pdf_writer = PyPDF2.PdfWriter()  
  
        # Merge all pages with the background  
        for page_number in range(len(content_pdf.pages)):  
            content_page = content_pdf.pages[page_number]  
            background_page = copy(background_pdf.pages[0])
            background_page.merge_page(content_page)  
            pdf_writer.add_page(background_page)  
  
        with open(final_pdf, 'wb') as output_file:  
            pdf_writer.write(output_file)  
  
    os.remove(temp_pdf) 
     
def export_images(df, full_folder_path, background_pdf_path):  
    if df.empty:  
        return {"error": "Please load a CSV file first."}  
  
    # Prepare order data for the pick_list.pdf  
    order_data = [['Order - Number', 'Item - Qty', 'Item - SKU', 'Item - Options', 'Item - Name', 'Custom - Field 3']]  
    for _, row in df.iterrows():  
        order_data.append([  
            row['Order - Number'],  
            row['Item - Qty'],  
            row['Item - SKU'],  
            row['Item - Options'],  
            row['Item - Name'],  
            row['Custom - Field 3']  
        ])  
  
    # Create the pick_list.pdf  
    create_pick_list_pdf(order_data, background_pdf_path, match_sku_to_item_description)  
  
  
    return {"message": f"PDF exported to {full_folder_path}!"}  
  
csv_file_path = "C:\\Users\\userID\\Desktop\\pick list\\test_list.csv"  
output_folder_path = "C:\\Users\\userID\\Downloads\\"  
background_pdf_path = "C:\\Users\userID\\Desktop\pick list\\GUIT\\background\\pick_list.pdf" 
  
df = pd.read_csv(csv_file_path)  
result = export_images(df, output_folder_path, background_pdf_path)  
print(result["message"]) 