import os
import fitz  # PyMuPDF
import shutil
import logging
import re
from PIL import Image
from docx import Document
from docx.shared import Inches
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters
from telegram.error import TimedOut

# === CONFIG ===
BOT_TOKEN = "7869425471:AAE7si4u_4jDqWRJPJ_VZwoO_wIC5pHK8_0" 

# === Logging ===
logging.basicConfig(level=logging.INFO)

# === Convert PDF to JPEG Images ===
def convert_pdf_to_images(pdf_path, output_folder="images"):
    os.makedirs(output_folder, exist_ok=True)
    doc = fitz.open(pdf_path)

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        mat = fitz.Matrix(2, 2) 
        pix = page.get_pixmap(matrix=mat, alpha=False)

        image_path = os.path.join(output_folder, f"page_{page_num+1}.jpg")
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Optional resize to reduce image size and speed up Word creation
        img = img.resize((int(pix.width * 0.7), int(pix.height * 0.7)))
        img.save(image_path, "JPEG", quality=75, optimize=True)
    doc.close()

# === Add Images to Word Document ===
def create_word_from_images(image_folder, output_file):
    def natural_sort_key(s):
        return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

    doc = Document()
    for img in sorted(os.listdir(image_folder), key=natural_sort_key):
        if img.lower().endswith((".jpg", ".jpeg")):
            doc.add_picture(os.path.join(image_folder, img), width=Inches(5.5))
            doc.add_page_break()
    doc.save(output_file)

# === Generate Unique Output Filename ===
def create_unique_filename(pdf_filename):
    base = re.sub(r"\.pdf$", "", pdf_filename, flags=re.IGNORECASE)
    filename = f"{base}_converted.docx"
    count = 1
    while os.path.exists(filename):
        filename = f"{base}_converted_{count}.docx"
        count += 1
    return filename

# === Telegram Handlers ===
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📥 Send me a PDF and I’ll convert it to a Word file with each page as an image.")

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = await update.message.reply_text("🔄 Downloading your PDF...")
    file = await update.message.document.get_file()
    file_path = await file.download_to_drive()

    pdf_filename = update.message.document.file_name or "document.pdf"
    output_filename = create_unique_filename(pdf_filename)

    try:
        await msg.edit_text("📄 Converting PDF to images...")
        convert_pdf_to_images(file_path)

        await msg.edit_text("🖼️ Creating Word document...")
        create_word_from_images("images", output_filename)

        await msg.edit_text("📤 Sending Word file...")
        await update.message.reply_document(document=open(output_filename, "rb"))

    except TimedOut:
        await update.message.reply_text("⏱️ Operation timed out. Try again with a smaller file.")
    except Exception as e:
        await update.message.reply_text(f"❌ Error: {e}")
    finally:
        for f in [output_filename, file_path]:
            if os.path.exists(f):
                os.remove(f)
        shutil.rmtree("images", ignore_errors=True)
        try:
            await msg.delete()
        except:
            pass

# === Run the Bot ===
def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
    app.run_polling()

if __name__ == "__main__":
    main()
