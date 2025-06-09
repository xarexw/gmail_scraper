import imaplib
import email
import os
from email.header import decode_header
from dotenv import load_dotenv
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# --- Налаштування ---
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# .env
load_dotenv()
GDRIVE_FOLDER_ID = os.getenv("GDRIVE_FOLDER_ID")
EMAIL_ACCOUNT = os.getenv("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
IMAP_SERVER = os.getenv("IMAP_SERVER")
DOWNLOAD_FOLDER = os.getenv("DOWNLOAD_FOLDER")

# Labels > names
# Якщо лист із міткою "partnerA" — вкладення буде "partnerA_price.xlsx"
labels_to_filename = {
    "Bestparts": "bestparts_price.xlsx",
    "Eminia New": "eminia_new_price.xls",
    "Masterteile": "masterteile_price.xls", #в них старий формат
    "MaxParts": "maxparts_price.xlsx",
    "Mtechno": "mtechno_price.xlsx",
    "Sprint": "sprint_price.xlsx",
    "Sprint All": "sprint_all_price.xlsx",
    "Syndicar": "syndicar_price.zip",
    "Ukrauto": "ukrauto_price.xlsx",
    "Usamotors": "usamotors_price.xlsx",
}

# Checklist of labels
labels_to_check = list(labels_to_filename.keys())

def get_drive_service():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    else:
        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return build("drive", "v3", credentials=creds)


def upload_to_drive(file_path, filename):
    service = get_drive_service()

    # Пошук файлу з таким іменем у папці GDRIVE_FOLDER_ID
    query = f"name = '{filename}' and '{GDRIVE_FOLDER_ID}' in parents and trashed = false"
    results = service.files().list(q=query, spaces='drive', fields="files(id, name)").execute()
    files = results.get('files', [])

    media = MediaFileUpload(file_path, resumable=True)

    if files:
        # Файл існує — оновлюємо
        file_id = files[0]['id']
        updated_file = service.files().update(
            fileId=file_id,
            media_body=media
        ).execute()
        print(f"Updated file on Drive: {filename} (id={file_id})")
    else:
        # Файл не знайдено — створюємо новий
        file_metadata = {
            "name": filename,
            "parents": [GDRIVE_FOLDER_ID] if GDRIVE_FOLDER_ID else []
        }
        created_file = service.files().create(
            body=file_metadata, media_body=media, fields="id"
        ).execute()
        print(f"Uploaded new file to Drive: {filename} (id={created_file.get('id')})")


# Decode utf-08
# Декодує у форматі MIME-Header (наприклад =?UTF-8?Q?...?=).


def decode_mime_words(s):
    decoded_fragments = decode_header(s)
    pieces = []
    for fragment, encoding in decoded_fragments:
        if isinstance(fragment, bytes):
            if encoding:
                pieces.append(fragment.decode(encoding, errors="ignore"))
            else:
                pieces.append(fragment.decode(errors="ignore"))
        else:
            pieces.append(fragment)
    return "".join(pieces)


# Link gmail with imap
try:
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
except imaplib.IMAP4.error as e:
    print(f"Login error: {e}")
    exit(1)

# Check labels separately
# Gmail в IMAP-інтерфейсі бачить мітки як “папки”.
# Щоб зайти в мітку, її треба брати в подвійні лапки, якщо є пробіли/спеціальні символи.
for label in labels_to_check:
    try:
        status, _ = mail.select(f'"{label}"')  # відкриваємо "папку" з назвою мітки
    except imaplib.IMAP4.error as e:
        print(f"Can't open label '{label}': {e}")
        continue

    if status != "OK":
        # Якщо немає такої “папки” (мітки) - переходимо далі
        print(f"Label '{label}' not found or have restrict access")
        continue

    # Шукаємо непрочитані (UNSEEN) листи в цій “папці” (мітці)
    status, data = mail.search(None, "UNSEEN")
    if status != "OK":
        print(f"can't search in label '{label}'.")
        continue

    email_ids = data[0].split()
    if not email_ids:
        # Немає нових листів з цією міткою
        continue

    for email_id in email_ids:
        # Завантажуємо сам лист
        status, msg_data = mail.fetch(email_id, "(RFC822)")
        if status != "OK":
            print(f"can't download file {email_id} in label '{label}'.")
            continue

        # Розбираємо байти в email.message.Message об’єкт
        raw_email = msg_data[0][1]
        msg = email.message_from_bytes(raw_email)

        # Доп. інформація про лист (корисно для дебагу)
        subject = msg.get("Subject", "")
        subject = decode_mime_words(subject)
        print(f"Checking mail with id={email_id.decode()} subject: \"{subject}\", label=\"{label}\"")

        used_filenames = set()
        attachments = []
        # --- Проходимо по всіх частинах листа, шукаємо прикріплення ---
        for part in msg.walk():
            # Якщо це контейнер multipart — пропускаємо
            if part.get_content_maintype() == "multipart":
                continue

            # Content-Disposition з вкладеннями містить “attachment”
            content_disposition = part.get("Content-Disposition", "")
            if not content_disposition or "attachment" not in content_disposition.casefold():
                continue

            # дізнаємося справжній content_type (наприклад, "application/pdf" або
            # "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            content_type = part.get_content_type()
            if not content_type.startswith("application"):
                # якщо це не "application/..." — пропускаємо
                continue

            # витягуємо розширення оригінальної назви файлу (filename ще не отримали,
            # але знаємо, що воно вказане в заголовку Content-Disposition або Content-Type)
            # спочатку потрібно отримати ім’я:
            orig_filename = part.get_filename()
            if not orig_filename:
                orig_filename = "attachment"
            else:
                orig_filename = decode_mime_words(orig_filename)

            ext = os.path.splitext(orig_filename)[1].lower()
            if ext not in [".xlsx", ".xls", ".zip"]:
                # якщо розширення не з-поміж тих - пропускаємо
                continue
            # Завантажуємо бінарні дані
            try:
                file_data = part.get_payload(decode=True)
            except Exception as e:
                print(f"Failed to decode attachment from message id={email_id.decode()}: {e}")
                continue

            # --- Перейменування згідно з міткою ---
            unified_name = labels_to_filename.get(label)
            if not unified_name:
                # Якщо з якоїсь причини не знайшлося ім'я для цієї мітки
                print(f"not found name of label '{label}', use original name {orig_filename}")
                unified_name = orig_filename  # у такому випадку лишаємо оригінал

            # Унікалізуємо ім’я файлу (+ _1)
            base_name, ext = os.path.splitext(unified_name)
            save_path = os.path.join(DOWNLOAD_FOLDER, unified_name)

            # --- Збереження файлів ---
            used = set()
            os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
            for file_data, unified_name in attachments:
                base, ext = os.path.splitext(unified_name)
                name = unified_name
                i = 1
                while name in used:
                    name = f"{base}_{i}{ext}"
                    i += 1
                used.add(name)

            # Гарантуємо, що папка для зберігання існує
            os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
            # save_path = os.path.join(DOWNLOAD_FOLDER, unified_name)

            # Записуємо у файл (перезаписуємо старий, якщо був)
            try:
                with open(save_path, "wb") as f:
                    f.write(file_data)
                print(f"save as \"{save_path}\" in \"{DOWNLOAD_FOLDER}\"")
                upload_to_drive(save_path, os.path.basename(save_path))
            except Exception as e:
                print(f"error with writing file \"{save_path}\": {e}")

        # --- Додаємо прапорець SEEN, щоби лист більше не потрапляв у UNSEEN ---
        mail.store(email_id, "+FLAGS", "\\Seen")

    try:
        mail.close()
    except Exception as e:
        print(f"Warning: could not close mailbox '{label}': {e}")

# Завершуємо сесію
mail.logout()
print("IMAP script done")
