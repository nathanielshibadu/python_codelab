import logging
import os
import pandas as pd
from functions import read_student_data, generate_emails_for_students

# Ensure the 'logs' directory exists
log_directory = 'logs'
if not os.path.exists(log_directory):
    os.makedirs(log_directory)

# Set up logging
logging.basicConfig(
    filename=os.path.join(log_directory, 'computations.log'),
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def main():
    # File path for the input Excel file
    file_path = 'data/Test Files.xlsx'

    # Step 1: Read the data from Excel file
    logging.info("Reading student data from Excel file.")
    try:
        # Read data separately from both sheets
        df_file_a, df_file_b = read_student_data(file_path)
        logging.info("Successfully read student data from File_A and File_B sheets.")
    except Exception as e:
        logging.error(f"Error reading student data: {e}")
        return

    # Step 2: Generate email addresses for File_A sheet
    logging.info("Generating email addresses for students in File_A sheet.")
    try:
        df_file_a_with_emails = generate_emails_for_students(df_file_a)
        logging.info("Successfully generated email addresses for students in File_A.")
    except Exception as e:
        logging.error(f"Error generating email addresses for File_A: {e}")
        return

    # Step 3: Generate email addresses for File_B sheet
    logging.info("Generating email addresses for students in File_B sheet.")
    try:
        df_file_b_with_emails = generate_emails_for_students(df_file_b)
        logging.info("Successfully generated email addresses for students in File_B.")
    except Exception as e:
        logging.error(f"Error generating email addresses for File_B: {e}")
        return

    # Step 4: Save the updated data to new Excel files
    output_file_a_path = 'data/Student_Emails_File_A.xlsx'
    output_file_b_path = 'data/Student_Emails_File_B.xlsx'

    try:
        df_file_a_with_emails.to_excel(output_file_a_path, index=False)
        df_file_b_with_emails.to_excel(output_file_b_path, index=False)
        logging.info(f"Emails generated and saved to {output_file_a_path} and {output_file_b_path}.")
    except Exception as e:
        logging.error(f"Error saving data to Excel files: {e}")
        return

    # Step 5: Save files to TSV and CSV formats
    csv_output_path_a = 'data/Student_Emails_File_A.csv'
    tsv_output_path_a = 'data/Student_Emails_File_A.tsv'
    csv_output_path_b = 'data/Student_Emails_File_B.csv'
    tsv_output_path_b = 'data/Student_Emails_File_B.tsv'

    try:
        logging.info(f"Saving data to CSV format at {csv_output_path_a} and {csv_output_path_b}.")
        df_file_a_with_emails.to_csv(csv_output_path_a, index=False)
        df_file_b_with_emails.to_csv(csv_output_path_b, index=False)

        logging.info(f"Saving data to TSV format at {tsv_output_path_a} and {tsv_output_path_b}.")
        df_file_a_with_emails.to_csv(tsv_output_path_a, sep='\t', index=False)
        df_file_b_with_emails.to_csv(tsv_output_path_b, sep='\t', index=False)

        logging.info("Data successfully saved in TSV and CSV formats.")
        print("Data saved in TSV and CSV formats.")

    except Exception as e:
        logging.error(f"Error saving data to CSV/TSV formats: {e}")
        return

if __name__ == "__main__":
    main()
