import pandas as pd
import re
import json
from sentence_transformers import SentenceTransformer
from scipy.spatial.distance import cosine
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
import os

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/drive.file']

def authenticate():
    """Authenticates and returns Google Drive API credentials."""
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

def upload_file_to_drive(file_name, file_path, mime_type):
    """Uploads a file to Google Drive."""
    creds = authenticate()
    service = build('drive', 'v3', credentials=creds)
    file_metadata = {'name': file_name}
    media = MediaFileUpload(file_path, mimetype=mime_type)
    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f"File '{file_name}' uploaded with file ID: {file.get('id')}")


def read_student_data(file_path):
    try:
        sheet_dict = pd.read_excel(file_path, sheet_name=['File_A', 'File_B'])
        df_file_a = sheet_dict['File_A']
        df_file_b = sheet_dict['File_B']
        return df_file_a, df_file_b
    except Exception as e:
        raise FileNotFoundError(f"Error reading Excel file: {e}")

def generate_email(name, existing_emails):
    name_parts = re.split(r'\s+', re.sub(r'[^\w\s]', '', name.lower()))
    if len(name_parts) == 1:
        email_base = name_parts[0]
    else:
        email_base = name_parts[0][0] + name_parts[-1]

    email_base = email_base.strip()
    domain = '@gmail.com'

    email = email_base + domain
    counter = 1
    while email in existing_emails:
        email = f"{email_base}{counter}{domain}"
        counter += 1

    existing_emails.add(email)
    return email

def generate_emails_for_students(df):
    existing_emails = set()
    df['Email Address'] = df["Student Name"].apply(lambda name: generate_email(name, existing_emails))
    return df

def generate_gender_lists(df):
    male_students = df[df['Gender'] == 'M']['Student Name'].tolist()
    female_students = df[df['Gender'] == 'F']['Student Name'].tolist()
    return male_students, female_students

def find_special_characters(df):
    special_char_pattern = r'[^a-zA-Z\s]'  # Excludes commas
    special_char_students = df[df['Student Name'].str.contains(special_char_pattern)]['Student Name'].tolist()
    return special_char_students

def name_similarity_analysis(df):
    model = SentenceTransformer('LaBSE')

    male_names = df[df['Gender'] == 'M']['Student Name'].tolist()
    female_names = df[df['Gender'] == 'F']['Student Name'].tolist()

    male_embeddings = model.encode(male_names)
    female_embeddings = model.encode(female_names)

    similarity_results = []
    for i, male_name in enumerate(male_names):
        for j, female_name in enumerate(female_names):
            similarity = 1 - cosine(male_embeddings[i], female_embeddings[j])
            if similarity >= 0.5:
                similarity_results.append({
                    "male_name": male_name,
                    "female_name": female_name,
                    "similarity": float(similarity)
                })

    similarity_results.sort(key=lambda x: x['similarity'], reverse=True)

    with open('data/name_similarity_results.json', 'w') as f:
        json.dump(similarity_results, f, indent=2)

    return similarity_results
