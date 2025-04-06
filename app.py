import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
import os

# Load Excel file
df = pd.read_excel("demo.xlsx")  # Ensure columns: "Name", "Amount"

# Load certificate template
template_path = "Share Certificate.png"

# ✅ Load fonts (Ensure the .ttf files exist in the same directory)
name_font_path = os.path.abspath("PinyonScript-Regular.ttf")  # Font for Name
amount_font_path = os.path.abspath("DejaVuSans-Bold.ttf")  # Font for Amount

# Output folder
output_folder = "certificates_pdf"
os.makedirs(output_folder, exist_ok=True)

# Check if the fonts exist
if not os.path.exists(name_font_path):
    print(f"❌ Error: {name_font_path} not found! Please check the file path.")
    exit()

if not os.path.exists(amount_font_path):
    print(f"❌ Error: {amount_font_path} not found! Please check the file path.")
    exit()

# Load template image to get size
img_template = Image.open(template_path)
img_width, img_height = img_template.size  # Get image dimensions

# Define Y positions for Name & Amount
name_y = 520   # Adjust as needed
amount_y = 915  # Adjust as needed

# Define font sizes
name_font_size = 100   # Large size for name
amount_font_size = 35  # Slightly smaller for amount

# Load fonts
try:
    name_font = ImageFont.truetype(name_font_path, name_font_size)
    amount_font = ImageFont.truetype(amount_font_path, amount_font_size)
    print("✅ Fonts loaded successfully!")
except OSError:
    print("❌ Error: Could not load fonts. Check the font file paths.")
    exit()

# Generate certificates
for _, row in df.iterrows():
    name = row["Name"]
    amount = f"₹{row['Amount']}"  # Add ₹ sign before amount

    # Open a fresh copy of the template
    img = img_template.copy()
    draw = ImageDraw.Draw(img)

    # Get text width for centering
    name_bbox = draw.textbbox((0, 0), name, font=name_font)
    amount_bbox = draw.textbbox((0, 0), amount, font=amount_font)

    name_width = name_bbox[2] - name_bbox[0]
    amount_width = amount_bbox[2] - amount_bbox[0]

    name_x = (img_width - name_width) // 2  # Center the name
    amount_x = (img_width - amount_width) // 2  # Center the amount

    # Add text to image
    draw.text((name_x, name_y), name, font=name_font, fill="black")
    draw.text((amount_x, amount_y), amount, font=amount_font, fill="black")

    # Save image temporarily
    temp_img_path = f"{output_folder}/{name}.png"
    img.save(temp_img_path)

    # ✅ FIX: Use `reportlab` to create PDF with exact image size
    pdf_output_path = f"{output_folder}/{name}.pdf"
    c = canvas.Canvas(pdf_output_path, pagesize=(img_width, img_height))
    c.drawImage(temp_img_path, 0, 0, img_width, img_height)  # Draw full-size image
    c.save()

    # Remove temporary image
    os.remove(temp_img_path)

    print(f"Generated PDF: {pdf_output_path}")

print("✅ All Certificates Generated with Different Fonts for Name & Amount!")
