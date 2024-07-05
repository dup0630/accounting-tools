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
        email_msg = extract_msg.Message(email)
        if email_msg.attachments:
            statements = email_msg.attachments
            for statement in statements:
                file_name = str(n) + statement.longFilename
                statement_path = os.path.join(output_folder, file_name)
                with open(statement_path, 'wb') as file:
                    file.write(statement.data)
                n += 1
                print(f"Extracted: {statement.longFilename}")
        else:
            print(f"The message '{email}' does not have any attachments.")



msg_file_path = input("Enter .MSG file name: ")
dir = 'Statements'
extract_attachments(msg_file_path, dir)