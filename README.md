# Accounting Tools
This repository contains some tools that I developed while working in an accounting setting. These tools are meant to automate repetitive tasks and reduce workload.

## Contents
### `statement_verification_template`
This folder contains VBA scripts for a macro-enabled Excel template that automates the loading and exploration of transaction records. Specifically, given an account statement (with outstanding invoices) and an invoice database (from either Workday or Epicor services), the template provides tools for automatically formatting the statement, importing the relevant local data, and comparing payment status or flagging missing invoices.

### `email_attachment_extractor`
This Python script processes an Outlook email containing multiple attached .msg files. It opens each attached message, extracts any attachments within them, and saves those files to a specified directory. The script includes a simple user interface to guide the user through the process.
