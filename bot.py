import os
from pdf2image import convert_from_path
import shutil
import logging
import re
from PIL import Image
from docx import Document
from docx.shared import Inches
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters
from telegram.error import TimedOut

# === Logging ===
logging.basicConfig(level=logging.INFO)

# === Telegram Bot Token from ENV ===
BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")

# === Convert PDF to JPEG Images ===
def convert_pdf_to_images(pdf_path, output_folder="images"):
    os.makedirs(output_folder, exist_ok=True)
    images = convert_from_path(pdf_path, dpi=150)
    for idx, img in enumerate(images, start=1):
        width, height = img.size
        img = img.resize((int(width * 0.7), int(height * 0.7)))
        image_path = os.path.join(output_folder, f"page_{idx}.jpg")
        img.save(image_path, "JPEG", quality=75, optimize=True)

def create_word_from_images(image_folder, output_file):
    def natural_sort_key(s):
        return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

    doc = Document()
    for img in sorted(os.listdir(image_folder), key=natural_sort_key):
        if img.lower().endswith((".jpg", ".jpeg")):
            doc.add_picture(os.path.join(image_folder, img), width=Inches(5.5))
            doc.add_page_break()
    doc.save(output_file)

def create_unique_filename(pdf_filename):
    base = re.sub(r"\.pdf$", "", pdf_filename, flags=re.IGNORECASE)
    filename = f"{base}_converted.docx"
    count = 1
    while os.path.exists(filename):
        filename = f"{base}_converted_{count}.docx"
        count += 1
    return filename

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

def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.PDF, handle_pdf))
    app.run_polling()

if __name__ == "__main__":
    main()
