# Clear all variables from the current Python namespace
globals().clear()

import os
import re
import json
from fpdf import FPDF

# Read firm marketing plan
path = "./output-dashboard/"
file = "FR.LST"
input_file = path + file
output_file = "fr.pdf"

# def text_to_pdf(input_file, output_file):
# Initialize the PDF
pdf = FPDF(orientation='L')  # 'L' is for landscape
pdf.set_auto_page_break(auto=False)  # Disable auto page breaks
pdf.set_font("Courier", size=10)    # Use monospace font (Courier)

# Set custom margins
pdf.set_margins(left=10, top=10, right=10)  # Adjust margins (left, top, right)


# Read the text file
with open(input_file, 'r') as file:
    lines = file.readlines()

# Variables to track page breaks
pdf.add_page()
for line in lines:
    line = line.rstrip("\n")  # Strip trailing newline
    
    # If line starts with "1", add a new page and skip the "1"
    if line.startswith("1"):
        pdf.add_page()
        line = line[1:].lstrip()  # Remove the "1" and leading spaces

    # Add the line to the PDF
    pdf.cell(0, 3, line, ln=True)

# Save the PDF to the output file
pdf.output(output_file)

