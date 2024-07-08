import os
import extract_msg

def extract_msg_from_msg(msg_file_path, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    with extract_msg.openMsg(msg_file_path) as msg:
        emails = msg.attachments
        for email in emails:
            email.save(customPath = output_dir, extractEmbedded = True)

def extract_statements(input_dir, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    for filename in os.listdir(input_dir):
        if filename.endswith('.msg'):
            path = os.path.join(input_dir, filename)
            with extract_msg.openMsg(path) as msg:
                msg.saveAttachments(customPath=output_dir)
        print(f"Extracted {filename}")


source = 'Source.msg'
dir1 = 'Emails'
dir2 = 'Statements'

# Process the .MSG files
extract_msg_from_msg(source, dir1)
extract_statements(dir1, dir2)