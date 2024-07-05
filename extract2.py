import extract_msg
import os

def extract_attachments(msg_file_path, output_folder):
    # Ensure the output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Open the .msg file
    msg = extract_msg.Message(msg_file_path)
    vendor_emails = msg.attachments

    # Process each attachment
    n = 1
    for email in vendor_emails:
        file_name = str(n) + email.longFilename
        email_path = os.path.join(output_folder, file_name)
        with open(email_path, 'wb') as file:
            file.write(email.data)
        n += 1
        print(f"Extracted: {email.longFilename}")


msg_file_path = input("Enter .MSG file name: ")
dir = 'Emails'
extract_attachments(msg_file_path, dir)