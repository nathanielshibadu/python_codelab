from functions import read_student_data, generate_emails_for_students

# File path for the input Excel file
file_path = 'data/Test Files.xlsx'

# Step 1: Read the data
df = read_student_data(file_path)

# Step 2: Generate email addresses
df_with_emails = generate_emails_for_students(df)

# Step 3: Save the updated data to a new Excel file
output_file_path = 'data/Student_Emails.xlsx'
df_with_emails.to_excel(output_file_path, index=False)
print(f"Emails generated and saved to {output_file_path}")
