import os
import extract_msg

# Function to extract all attachments from a single .MSG file
def extract_attachments(msg_file_path, output_folder):
    msg = extract_msg.Message(msg_file_path)
    attachments = msg.attachments
    n = 1
    for attachment in attachments:
        attachment_path = os.path.join(output_folder, attachment.longFilename)
        with open(attachment_path, 'wb') as file:
            file.write(attachment.data)
        print(f"Extracted: {attachment.longFilename}")
        n += 1

# Function to process all .MSG files in a directory
def process_msg_files(input_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    for filename in os.listdir(input_folder):
        if filename.endswith('.msg'):
            msg_file_path = os.path.join(input_folder, filename)
            extract_attachments(msg_file_path, output_folder)

# Input and output directories
email_path = input(".MSG file name: ")
input = 'Emails'
output = 'Statements'

# Process the .MSG files
extract_attachments(email_path, input)
process_msg_files(input, output)
