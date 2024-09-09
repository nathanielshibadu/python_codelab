import pandas as pd
import re

# Function to read student data from an Excel file
def read_student_data(file_path):
    df = pd.read_excel(file_path)
    return df

# Function to generate email address based on student name
def generate_email(name, existing_emails):
    # Remove special characters from the name
    name_parts = re.split(r'\s+', re.sub(r'[^\w\s]', '', name.lower()))
    # Use the first letter of the first name and the last name for the email
    if len(name_parts) == 1:
        email_base = name_parts[0]
    else:
        email_base = name_parts[0][0] + name_parts[-1]

    email_base = email_base.strip()
    domain = '@gmail.com'

    # Ensure email is unique
    email = email_base + domain
    counter = 1
    while email in existing_emails:
        email = f"{email_base}{counter}{domain}"
        counter += 1

    existing_emails.add(email)
    return email

# Function to generate emails for all students
def generate_emails_for_students(df):
    existing_emails = set()
    df['email'] = df["Student Name"].apply(lambda name: generate_email(name, existing_emails))
    return df
