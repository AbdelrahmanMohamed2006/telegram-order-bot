# ================================
# Telegram Bot with Webhook (Render Ready)
# ================================

import os
import logging
import re
from datetime import datetime

from telegram import Update
from telegram.request import HTTPXRequest
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

from extractor import extract_order_data
from excel_generator import create_excel

# Logging
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Bot Token
TOKEN = os.getenv("TELEGRAM_TOKEN")
if not TOKEN:
    raise ValueError("❌ TELEGRAM_TOKEN is missing in environment variables")

# Folder for temp files
DOWNLOAD_FOLDER = "temp_orders"
if not os.path.exists(DOWNLOAD_FOLDER):
    os.makedirs(DOWNLOAD_FOLDER)

# Store each user's files
user_files = {}

# =====================================================
# Handlers
# =====================================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 أهلاً بك!\n"
        "أرسل ملفات Word (DOCX)، واحدة تلو الأخرى.\n"
        "وعند الانتهاء أرسل: /done"
    )


async def handle_docx(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id

    if update.message.document.mime_type != \
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document":

        await update.message.reply_text("⚠️ أرسل ملفات DOCX فقط.")
        return

    file_name = update.message.document.file_name
    new_file = await update.message.document.get_file()

    safe_name = re.sub(r"[^a-zA-Z0-9._]", "_", file_name)
    save_path = os.path.join(DOWNLOAD_FOLDER, f"{user_id}_{safe_name}")

    await new_file.download_to_drive(save_path)

    if user_id not in user_files:
        user_files[user_id] = []
    user_files[user_id].append(save_path)

    await update.message.reply_text(
        f"📄 تم استلام الملف: **{file_name}**\n"
        f"العدد الإجمالي حتى الآن: **{len(user_files[user_id])}**",
        parse_mode="Markdown"
    )


async def process_and_send(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    chat_id = update.effective_chat.id

    if user_id not in user_files or len(user_files[user_id]) == 0:
        await update.message.reply_text("❌ لم تستلم أي ملفات.")
        return

    files = user_files[user_id]
    extracted_data = []

    await context.bot.send_message(
        chat_id, f"⏳ جاري معالجة {len(files)} ملف..."
    )

    for path in files:
        data = extract_order_data(path)
        if data and data.get("رقم_الأمر"):
            extracted_data.append(data)

    if len(extracted_data) == 0:
        await context.bot.send_message(chat_id, "❌ لا توجد بيانات صالحة.")
    else:
        excel_name = f"Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        excel_path = os.path.join(DOWNLOAD_FOLDER, excel_name)

        create_excel(extracted_data, excel_path)

        with open(excel_path, "rb") as f:
            await context.bot.send_document(
                chat_id,
                document=f,
                caption=f"✅ تم إنشاء تقرير Excel يحتوي على {len(extracted_data)} صف."
            )

        os.remove(excel_path)

    # Cleanup
    for path in files:
        if os.path.exists(path):
            os.remove(path)

    del user_files[user_id]

    await context.bot.send_message(chat_id, "🗑️ تم تنظيف الملفات المؤقتة.")


# =====================================================
# Webhook (Render hosting)
# =====================================================

def main():
    PORT = int(os.environ.get("PORT", 8443))
    WEBHOOK_URL = os.getenv("WEBHOOK_URL")

    if not WEBHOOK_URL:
        raise ValueError("❌ WEBHOOK_URL must be defined on Render")

    request_obj = HTTPXRequest(read_timeout=30)

    app = (
        Application.builder()
        .token(TOKEN)
        .request(request_obj)
        .build()
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("done", process_and_send))
    app.add_handler(MessageHandler(filters.Document.DOCX, handle_docx))

    logger.info("🚀 Starting bot with webhook...")

    app.run_webhook(
        listen="0.0.0.0",
        port=PORT,
        url_path=TOKEN,
        webhook_url=f"{WEBHOOK_URL}/{TOKEN}"
    )


if __name__ == "__main__":
    main()
