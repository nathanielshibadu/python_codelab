import pandas as pd
import re
from sentence_transformers import SentenceTransformer
from scipy.spatial.distance import cosine
import json

def read_student_data(file_path):
    try:
        sheet_dict = pd.read_excel(file_path, sheet_name=['File_A', 'File_B'])
        df_file_a = sheet_dict['File_A']
        df_file_b = sheet_dict['File_B']
        return df_file_a, df_file_b
    except Exception as e:
        raise FileNotFoundError(f"Error reading Excel file: {e}")

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

def name_similarity_analysis(df):
    # Initialize LaBSE model
    model = SentenceTransformer('LaBSE')

    # Separate male and female names
    male_names = df[df['Gender'] == 'M']['Student Name'].tolist()
    female_names = df[df['Gender'] == 'F']['Student Name'].tolist()

    # Get embeddings for all names
    male_embeddings = model.encode(male_names)
    female_embeddings = model.encode(female_names)

    # Calculate similarity
    similarity_results = []
    for i, male_name in enumerate(male_names):
        for j, female_name in enumerate(female_names):
            similarity = 1 - cosine(male_embeddings[i], female_embeddings[j])
            if similarity >= 0.5:  # Only include results with at least 50% similarity
                similarity_results.append({
                    "male_name": male_name,
                    "female_name": female_name,
                    "similarity": float(similarity)
                })

    # Sort results by similarity (descending order)
    similarity_results.sort(key=lambda x: x['similarity'], reverse=True)

    # Save results to JSON file
    with open('data/name_similarity_results.json', 'w') as f:
        json.dump(similarity_results, f, indent=2)

    return similarity_results

def save_to_json(df, file_path):
    try:
        df.to_json(file_path, orient='records', lines=False, indent=2)
        print(f"Data saved to JSON file: {file_path}")
    except Exception as e:
        print(f"Error saving data to JSON: {e}")

def save_to_jsonl(df, file_path, similarity_results):
    try:
        with open(file_path, 'w') as f:
            for _, row in df.iterrows():
                # Convert all data in the row to string if needed
                json_record = {
                    "id": str(row.name),
                    "student_number": str(row['Student Number']),
                    "additional_details": [
                        {
                            "dob": str(row['DoB']) if pd.notnull(row['DoB']) else None,
                            "gender": row['Gender'].lower(),
                            "special_character": ["yes" if re.search(r'[^a-zA-Z\s,]', row['Student Name']) else "no"],
                            "name_similar": ["yes" if any(result['similarity'] >= 0.5 for result in similarity_results if
                                                          result['male_name'] == row['Student Name'] or result['female_name'] == row['Student Name']) else "no"]
                        }
                    ]
                }
                # f.write(json.dumps(json_record) + '\n')
                f.write(json.dumps(json_record, indent=4) + '\n')
        print(f"Data saved to JSONL file: {file_path}")
    except Exception as e:
        print(f"Error saving data to JSONL: {e}")
