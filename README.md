Certificate Generator from Excel (Python)

This Python script reads participant data from an Excel file and automatically generates personalized Word (.docx) certificates using a predefined certificate template.

Each row in the Excel file contains the participant’s name, event, and date, which are inserted into the certificate template to create individual certificates.

Features

Reads data from an Excel file

Uses a Word certificate template

Automatically fills in:

Participant name

Event name

Event date

Generates a separate .docx certificate for each entry

Fast and scalable for large lists

Input Format (Excel)

The Excel file should contain the following columns:

Name	Event	Date
John Doe	Data Science Workshop	12 June 2025
Jane Smith	Python Training	15 June 2025
Output

Individual Word certificate files (.docx)

File names can be customized (e.g. Certificate_John_Doe.docx)

How It Works

Prepare an Excel file with participant data.

Create a Word certificate template with placeholders.

Run the Python script.

Personalized certificates are generated automatically.
