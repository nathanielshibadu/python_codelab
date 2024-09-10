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

    # Step 1: Read the data
    logging.info("Reading student data from Excel file.")
    try:
        df = read_student_data(file_path)
        logging.info("Successfully read student data.")
    except Exception as e:
        logging.error(f"Error reading student data: {e}")
        return

    # Step 2: Generate email addresses
    logging.info("Generating email addresses for students.")
    try:
        df_with_emails = generate_emails_for_students(df)
        logging.info("Successfully generated email addresses for students.")
    except Exception as e:
        logging.error(f"Error generating email addresses: {e}")
        return

    # Step 3: Save the updated data to a new Excel file
    output_file_path = 'data/Student_Emails.xlsx'
    try:
        df_with_emails.to_excel(output_file_path, index=False)
        logging.info(f"Emails generated and saved to {output_file_path}.")
    except Exception as e:
        logging.error(f"Error saving data to Excel file: {e}")
        return

    # Step 4: Save files to TSV and CSV formats
    csv_output_path = 'data/Student_Emails.csv'
    tsv_output_path = 'data/Student_Emails.tsv'

    try:
        logging.info(f"Saving data to CSV format at {csv_output_path}.")
        df_with_emails.to_csv(csv_output_path, index=False)

        logging.info(f"Saving data to TSV format at {tsv_output_path}.")
        df_with_emails.to_csv(tsv_output_path, sep='\t', index=False)

        logging.info("Data successfully saved in TSV and CSV formats.")
        print("Data saved in TSV and CSV formats.")

    except Exception as e:
        logging.error(f"Error saving data to CSV/TSV formats: {e}")
        return

if __name__ == "__main__":
    main()
