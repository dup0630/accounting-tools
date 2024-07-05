import os
import extract_msg

# Function to extract all attachments from a single .MSG file
def extract_attachments_from_msg(msg_file_path, output_folder):
    msg = extract_msg.Message(msg_file_path)
    attachments = msg.attachments

    for attachment in attachments:
        attachment_path = os.path.join(output_folder, attachment.longFilename)
        with open(attachment_path, 'wb') as file:
            file.write(attachment.data)
        print(f"Extracted: {attachment.longFilename}")

# Function to process all .MSG files in a directory
def process_msg_files_in_directory(directory, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    for filename in os.listdir(directory):
        if filename.endswith('.msg'):
            msg_file_path = os.path.join(directory, filename)
            extract_attachments_from_msg(msg_file_path, output_folder)

# Input and output directories
input_directory = 'Emails'
output_directory = 'Statements'

# Process the .MSG files
process_msg_files_in_directory(input_directory, output_directory)
