import openpyxl
from PIL import Image, ImageDraw, ImageFont

# Load the workbook
workbook = openpyxl.load_workbook('detail.xlsx')

# Access the active sheet
sheet = workbook.active

# Extract the names to a list
names = []
for row in sheet.iter_rows(min_row=1, values_only=True):
    name = row[0]
    if name:
        names.append(name)

# Define a function to generate a certificate for a name
def generate_certificate(name):
    # Load the certificate template
    template = Image.open('certificate.jpg')

    # Create a drawing context
    draw = ImageDraw.Draw(template)

    # Load a font
    font = ImageFont.truetype('ShortBaby-Mg2w.ttf', 108)

    # Calculate the position to draw the name
    name_width, name_height = draw.textsize(name, font=font)
    name_x = (template.width - name_width) / 2
    name_y = 580

    # Draw the name on the certificate
    draw.text((name_x, name_y), name, font=font, fill=(0, 0, 0))

    # Save the certificate image
    template.save(f'{name}_certificate.jpg')

# Generate certificates for each name
for name in names:
    generate_certificate(name)
