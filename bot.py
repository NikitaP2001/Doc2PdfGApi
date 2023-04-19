# Import the necessary libraries
import os
import io
import random
import string
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from telegram import __version__ as TG_VER
from telegram import Update, Document, ForceReply
from telegram.ext import Application, ContextTypes, filters, Updater, CommandHandler, MessageHandler, CallbackContext
from ttoken import TOKEN
import tempfile


SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/documents']

FILE_PATH = 'example.doc'


def login():
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds


def upload_file_to_drive(file_path, credentials):
    try:
        service = build('drive', 'v3', credentials=credentials)

        file_metadata = {'name': os.path.basename(file_path)}
        media = MediaFileUpload(file_path)
        file = service.files().create(body=file_metadata, media_body=media,
                                      fields='id').execute()
        file_id = file.get('id')
        print('File ID: {}'.format(file_id))
        return file_id

    except HttpError as error:
        print('An error occurred while uploading the file to Drive: {}'.format(error))
        return None

def generate_random_string(length):
    return ''.join(random.choices(string.ascii_letters + string.digits, k=length))

def create_google_docs_file(creds):
    try:
        file_name = generate_random_string(10)
        file_metadata = {
            'name': file_name,
            'mimeType': 'application/vnd.google-apps.document'
        }
        service = build('drive', 'v3', credentials=creds)
        google_docs_file = service.files().create(
            body=file_metadata,
        ).execute()
        return google_docs_file.get('id')

    except HttpError as error:
        print('An error creating Editor file: {}'.format(error))
        return None

def upload_file(local_file_path, credentials):
    try:
        service = build('drive', 'v3', credentials=credentials)

        with io.open(local_file_path, 'rb') as f:
            document_content = f.read()

        media = MediaIoBaseUpload(io.BytesIO(document_content), mimetype= \
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document', \
                chunksize=1024 * 1024, resumable=True)

        file_metadata = {
            'mimeType': 'application/vnd.google-apps.document',
            'name': 'My Document'
        }
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        return file.get('id')

        return document_id
    except HttpError as error:
        print(f'An error occurred: {error}')
        return None


def convert_to_pdf(file_id, credentials, resname):
    try:
        service = build('drive', 'v3', credentials=credentials)

        export_path = resname.join('.pdf')
        response = service.files().export_media(fileId=file_id, mimeType='application/pdf').execute()

        with open(export_path, 'wb') as f:
            f.write(response)

        print('PDF has been saved to: {}'.format(export_path))

    except HttpError as error:
        print('An error occurred while converting the file to PDF: {}'.format(error))


def delete_file(file_id, creds):
    service = build('drive', 'v3', credentials=creds)
    try:
        service.files().delete(fileId=file_id).execute()
        print('File with ID {} has been deleted.'.format(file_id))
    except HttpError as error:
        print('An error occurred while deleting the file: {}'.format(error))

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    user = update.effective_user
    await update.message.reply_html(
        rf"Hi {user.mention_html()}! Send me doc, and i`ll return PDF.",
        reply_markup=ForceReply(selective=True),
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text("As been mentioned, send me a doc and i will return pdf")

async def download(update: Update, context: ContextTypes):
    export_path = None
    name = str(update.message.chat_id)
    doc_id = None
    try:
        file = await context.bot.get_file(update.message.document)
        await file.download_to_drive(name)
        creds = login()
        if creds and creds.valid:
            doc_id = upload_file(name, creds)
            if doc_id:
                convert_to_pdf(doc_id, creds, name)
                export_path = name.join('.pdf')
                document = open(export_path, 'rb')
                await context.bot.send_document(update.message.chat_id, document, filename='converted.pdf')

    except Exception as e:
        print(f'Error handling download: {e}')
        await update.message.reply_text('Convertation failed')
    finally:
        if doc_id:
            delete_file(doc_id, creds)
        if name:
            os.remove(name)
        if export_path:
            os.remove(export_path)


def main():
    app = Application.builder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))

    # on non command i.e message - echo the message on Telegram
    app.add_handler(MessageHandler(filters.Document.ALL, download))

    # Run the bot until the user presses Ctrl-C
    app.run_polling()
    exit(0)

    

if __name__ == '__main__':
    main()
