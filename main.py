import logging
import os
import pandas as pd
from functions import (
    read_student_data,
    generate_emails_for_students,
    generate_gender_lists,
    find_special_characters,
    name_similarity_analysis,
    save_to_json,
    save_to_jsonl
)

# Ensure the 'logs' and 'data' directories exist
for directory in ['logs', 'data']:
    if not os.path.exists(directory):
        os.makedirs(directory)

# Set up logging
logging.basicConfig(
    filename=os.path.join('logs', 'computations.log'),
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def main():
    # File path for the input Excel file
    file_path = 'data/Test Files.xlsx'

    # Step 1: Read the data from Excel file
    logging.info("Reading student data from Excel file.")
    try:
        df_file_a, df_file_b = read_student_data(file_path)
        logging.info("Successfully read student data from File_A and File_B sheets.")
    except Exception as e:
        logging.error(f"Error reading student data: {e}")
        return

    combined_df = pd.DataFrame()  # Initialize an empty DataFrame for combined data

    # Step 2: Process each sheet separately
    for sheet_name, df in [("File_A", df_file_a), ("File_B", df_file_b)]:
        logging.info(f"Processing {sheet_name} sheet.")
        try:
            # Generate email addresses
            df_with_emails = generate_emails_for_students(df)
            logging.info(f"Successfully generated email addresses for students in {sheet_name}.")

            # Generate separate lists of Male and Female students
            male_students, female_students = generate_gender_lists(df_with_emails)

            # Save separate lists of Male and Female students
            male_list_path = f'data/Male_Students_{sheet_name}.txt'
            female_list_path = f'data/Female_Students_{sheet_name}.txt'

            with open(male_list_path, 'w') as f:
                for student in male_students:
                    f.write(f"{student}\n")

            with open(female_list_path, 'w') as f:
                for student in female_students:
                    f.write(f"{student}\n")

            logging.info(f"{sheet_name}: Number of male students: {len(male_students)}")
            logging.info(f"{sheet_name}: Number of female students: {len(female_students)}")
            logging.info(f"Male students list saved to {male_list_path}")
            logging.info(f"Female students list saved to {female_list_path}")

            # Find students with special characters in their names
            special_char_students = find_special_characters(df_with_emails)
            logging.info(f"{sheet_name}: Students with special characters in their names: {special_char_students}")

            # Save the updated data to new Excel, CSV, and TSV files
            output_excel_path = f'data/Student_Emails_{sheet_name}.xlsx'
            output_csv_path = f'data/Student_Emails_{sheet_name}.csv'
            output_tsv_path = f'data/Student_Emails_{sheet_name}.tsv'

            df_with_emails.to_excel(output_excel_path, index=False)
            df_with_emails.to_csv(output_csv_path, index=False)
            df_with_emails.to_csv(output_tsv_path, sep='\t', index=False)

            logging.info(f"Data for {sheet_name} saved in Excel, CSV, and TSV formats.")

            # Append the processed data to the combined DataFrame
            combined_df = pd.concat([combined_df, df_with_emails], ignore_index=True)

        except Exception as e:
            logging.error(f"Error processing {sheet_name}: {e}")

    # Perform name similarity analysis on the combined data
    logging.info("Performing name similarity analysis.")
    try:
        similarity_results = name_similarity_analysis(combined_df)
        logging.info(f"Name similarity analysis complete. Results saved to 'data/name_similarity_results.json'.")
        logging.info(f"Number of name pairs with similarity >= 50%: {len(similarity_results)}")
    except Exception as e:
        logging.error(f"Error during name similarity analysis: {e}")
        return

    # Save combined data to JSON and JSONL formats
    json_output_path = 'data/Combined_Student_Data.json'
    jsonl_output_path = 'data/Combined_Student_Data.jsonl'

    save_to_json(combined_df, json_output_path)
    save_to_jsonl(combined_df, jsonl_output_path, similarity_results)

    print("Processing complete. Check the logs and output files in the 'data' directory for details.")

if __name__ == "__main__":
    main()
