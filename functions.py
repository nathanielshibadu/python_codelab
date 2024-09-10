import pandas as pd
import re


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
    special_char_pattern = r'[^a-zA-Z\s,]'
    special_char_students = df[df['Student Name'].str.contains(special_char_pattern)]['Student Name'].tolist()
    return special_char_students