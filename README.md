# Accounting Tools
This repository contains some tools that I developed while working in an accounting setting. These tools are meant to automate repetitive tasks and reduce workload.

## Contents
### `statement_verification_template`
This folder contains VBA scripts for a macro-enabled Excel template that automates the loading and exploration of transaction records. Specifically, given an account statement (with outstanding invoices) and an invoice database (from either Workday or Epicor services), the template provides tools for automatically formatting the statement, importing the relevant local data, and comparing payment status or flagging missing invoices.

### `email_attachment_extractor`
This Python script provides a simple GUI for extracting attachments from Outlook .msg files. It supports two modes:
- **Email Mode:** If the selected .msg file contains other .msg files as attachments (i.e., nested emails), the script extracts those first, then retrieves attachments from the nested emails.
- **Attachment Mode:** If the selected .msg file directly contains attachments (not other emails), the script extracts them directly.
All extracted files are saved to pre-defined directories. The user is guided through the process via a basic, intuitive interface.


