import imaplib
import email
import io
import pandas as pd
import threading
from telegram import Update
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters, CallbackContext
from dotenv import load_dotenv
import os
from email.header import decode_header

load_dotenv()

IMAP_SERVER = os.getenv('IMAP_SERVER')
EMAIL_ACCOUNT = os.getenv('EMAIL_ACCOUNT')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
SENDER_FILTERS = os.getenv('SENDER_FILTERS', '')
SENDER_FILTERS = [addr.strip().lower() for addr in SENDER_FILTERS.split(',') if addr.strip()]
TELEGRAM_TOKEN = os.getenv('TELEGRAM_TOKEN')
TELEGRAM_CHAT_ID = int(os.getenv('TELEGRAM_CHAT_ID'))

SEARCH_KEYWORD_FILE = "search_keyword.txt"

def save_keyword(word):
    with open(SEARCH_KEYWORD_FILE, "w", encoding="utf-8") as f:
        f.write(word.strip().lower())

def load_keyword():
    try:
        with open(SEARCH_KEYWORD_FILE, "r", encoding="utf-8") as f:
            return f.read().strip().lower()
    except FileNotFoundError:
        return ""

def decode_mime_words(s):
    if not s:
        return ''
    decoded = decode_header(s)
    return ''.join(
        str(part, charset or 'utf-8') if isinstance(part, bytes) else part
        for part, charset in decoded
    )

def send_telegram_file(bot, file_bytes, filename, caption=None):
    bot.send_document(
        chat_id=TELEGRAM_CHAT_ID,
        document=file_bytes,
        filename=filename,
        caption=caption if caption else ""
    )

def search_excel_for_keyword(file_data, keyword):
    try:
        excel_file = io.BytesIO(file_data)
        for sheet in pd.ExcelFile(excel_file).sheet_names:
            excel_file.seek(0)
            df = pd.read_excel(excel_file, sheet_name=sheet)
            if df.astype(str).apply(lambda col: col.str.lower().str.contains(keyword)).any().any():
                return True
        return False
    except Exception as e:
        return False

def connect_to_email(bot=None, verbose=False):
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
        mail.select('INBOX', readonly=True)
        if verbose and bot is not None:
            bot.send_message(chat_id=TELEGRAM_CHAT_ID, text="✅ Вход на почту успешно выполнен.")
        return mail
    except Exception as e:
        if verbose and bot is not None:
            bot.send_message(chat_id=TELEGRAM_CHAT_ID, text=f"❌ Ошибка входа на почту: {e}")
        raise

def check_emails(update=None, context=None, verbose=False, bot=None):
    if context is not None:
        bot = context.bot
    elif bot is None:
        updater = Updater(token=TELEGRAM_TOKEN, use_context=True)
        bot = updater.bot

    keyword = load_keyword()
    if not keyword:
        if verbose:
            bot.send_message(chat_id=TELEGRAM_CHAT_ID, text="❗ Ключевое слово для поиска не задано. Просто отправь его в чат.")
        return

    try:
        mail = connect_to_email(bot=bot, verbose=verbose)
    except Exception:
        return

    status, messages = mail.search(None, 'UNSEEN')
    if status != 'OK':
        if verbose:
            bot.send_message(chat_id=TELEGRAM_CHAT_ID, text="❌ Ошибка поиска писем.")
        mail.logout()
        return

    found_sender = False
    found_attachment = False
    found_keyword = False

    for num in messages[0].split():
        status, data = mail.fetch(num, '(RFC822)')
        if status != 'OK':
            continue

        msg = email.message_from_bytes(data[0][1])
        from_address = email.utils.parseaddr(msg.get("From"))[1]
        subject_raw = msg.get("Subject", "(без темы)")
        subject = decode_mime_words(subject_raw)

        if from_address.lower() in SENDER_FILTERS:
            found_sender = True
            if verbose:
                bot.send_message(chat_id=TELEGRAM_CHAT_ID,
                                text=f"📧 Письмо от {from_address} обнаружено (тема: {subject}).")
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue

                filename_raw = part.get_filename()
                filename = decode_mime_words(filename_raw)
                if filename and filename.lower().endswith('.xlsx'):
                    found_attachment = True
                    if verbose:
                        bot.send_message(chat_id=TELEGRAM_CHAT_ID,
                                        text=f"📎 Excel-вложение '{filename}' найдено.")
                    file_data = part.get_payload(decode=True)
                    if search_excel_for_keyword(file_data, keyword):
                        found_keyword = True
                        bot.send_message(
                            chat_id=TELEGRAM_CHAT_ID,
                            text=f"✅ Слово '{keyword}' найдено во вложении '{filename}'."
                        )
                        send_telegram_file(bot, io.BytesIO(file_data), filename,
                            caption=f"Вложение '{filename}', найдено слово '{keyword}'")
                    else:
                        if verbose:
                            bot.send_message(
                                chat_id=TELEGRAM_CHAT_ID,
                                text=f"❗ Слово '{keyword}' не найдено во вложении '{filename}'."
                            )
            if not found_attachment and verbose:
                bot.send_message(
                    chat_id=TELEGRAM_CHAT_ID,
                    text=f"❗ В письме от {from_address} нет Excel-вложений."
                )
    if not found_sender and verbose:
        bot.send_message(
            chat_id=TELEGRAM_CHAT_ID,
            text=f"❗ Нет новых писем от ({', '.join(SENDER_FILTERS)})."
        )
    mail.logout()

def set_keyword(update: Update, context: CallbackContext):
    keyword = update.message.text.strip().lower()
    if not keyword:
        update.message.reply_text("Отправь слово или фамилию для поиска.")
        return
    save_keyword(keyword)
    update.message.reply_text(f"Буду искать: '{keyword}'. Напиши /check чтобы проверить почту.")

def start(update: Update, context: CallbackContext):
    update.message.reply_text("Привет! Напиши слово или фамилию для поиска, затем команду /check.")

def check_command(update: Update, context: CallbackContext):
    check_emails(update, context, verbose=True)

def periodic_check(bot):
    check_emails(verbose=False, bot=bot)
    threading.Timer(600, periodic_check, args=(bot,)).start()  # 10 минут

def main():
    updater = Updater(token=TELEGRAM_TOKEN, use_context=True)
    dp = updater.dispatcher

    dp.add_handler(CommandHandler("start", start))
    dp.add_handler(CommandHandler("check", check_command))
    dp.add_handler(MessageHandler(Filters.text & ~Filters.command, set_keyword))

    threading.Timer(60, periodic_check, args=(updater.bot,)).start()  # Через минуту, потом каждые 10 минут

    updater.start_polling()
    updater.idle()

if __name__ == '__main__':
    main()
