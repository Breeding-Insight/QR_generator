import sys
import qrcode
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.shared import Inches
from io import BytesIO

def generate_qr_doc(input_txt_file, output_docx_file):
    def create_qr_image(data, qr_size=(300, 300), font_size=40):  # 1 x 1 inches, 300 DPI
        qr = qrcode.QRCode(
            version=1, 
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(data)
        qr.make(fit=True)

        qr_image = qr.make_image(fill='black', back_color='white')

        img = Image.new('RGB', (qr_size[0], qr_size[1] + 50), 'white')
        draw = ImageDraw.Draw(img)
        border_thickness = 5
        draw.rectangle([0, 0, img.size[0] - 1, img.size[1] - 1], outline='black', width=border_thickness)

        qr_image = qr_image.resize((int(qr_size[0] * 0.9), int(qr_size[0] * 0.9)), Image.LANCZOS)
        qr_position = ((qr_size[0] - qr_image.size[0]) // 2, 10) 
        img.paste(qr_image, qr_position)

        try:
            font = ImageFont.truetype("Arial.ttf", font_size)
        except IOError:
            font = ImageFont.load_default()

        text_bbox = draw.textbbox((0, 0), data, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        text_position = ((qr_size[0] - text_width) // 2, qr_position[1] + qr_image.size[1] + 5)

        draw.text(text_position, data, fill='black', font=font)

        img = img.resize((qr_size[0], qr_size[0]), Image.LANCZOS)

        return img

    def save_qr_to_docx(texts, output_file):
        doc = Document()
        for text in texts:
            print(f"Generating QR for: {text}")
            qr_image = create_qr_image(text)

            img_stream = BytesIO()
            qr_image.save(img_stream, format='PNG')
            img_stream.seek(0)

            # Adjust width to 1 inch in the DOCX
            doc.add_picture(img_stream, width=Inches(1))

        doc.save(output_file)

    with open(input_txt_file, 'r') as file:
        texts = file.read().splitlines()

    save_qr_to_docx(texts, output_docx_file)
    print(f"QR codes saved to {output_docx_file}.")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python3 qr_code_generator.py <input_txt_file> <output_docx_file>")
        sys.exit(1)

    input_txt_file = sys.argv[1]
    output_docx_file = sys.argv[2]

    generate_qr_doc(input_txt_file, output_docx_file)