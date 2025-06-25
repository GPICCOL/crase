# Clear all variables from the current Python namespace
globals().clear()

import os
import re
import json
from fpdf import FPDF
from PyPDF2 import PdfReader, PdfWriter

# Initialization parameters
total_firms = 6  # Total number of firms
path = "./output-dashboard/"
input_file = path + "FR.LST"
output_file = path + "facilitator.pdf"
firms = list(range(1, total_firms + 1))

# Function Definitions
# Creation of the full file in pdf (facilitator file)
def text_to_pdf(input_file, output_file):
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

# Creation of the individual firm file in pdf
def subset_pdf(input_pdf, output_pdf, pages):
  # Create a PdfReader object to read the PDF
  reader = PdfReader(input_pdf)
  
  # Create a PdfWriter object to write the new PDF
  writer = PdfWriter()
  
  # Loop through the list of pages to keep
  for page_num in pages:
      # Convert 1-based page numbers to 0-based index for PyPDF2
      if 0 <= page_num - 1 < len(reader.pages):  # Check that the page exists
          writer.add_page(reader.pages[page_num - 1])
  
  # Save the new PDF with selected pages
  with open(output_pdf, "wb") as output_file:
      writer.write(output_file)
  
  print(f"PDF with pages {pages_to_keep} saved as {output_pdf}")

# P&L and income statement extractor
def extract_pl_is(page, statement):
  # Define labels to extract
  if statement == "pl":
      labels = [
        "Food Sales", "Beverage Sales", "Food Cost", "Beverage cost",
        "Payroll", "Employee benefits", "Direct Operating",
        "Music & Entertainment", "Repairs & Maintenance",
        "Admin. & General", "Advertising & Promo.", "Utilities",
        "Franchise Fees", "Property Tax",
        "Rentals & misc.", "Liquor Lic. Fee",
        "Insurance", "Amortization", "Interest - Long term",
        "Depreciation", "Extraordinary Inc/Exp", "Interest - Short term",
        "Income Tax"
      ]
      
  elif statement == "is":
    labels = [
      "Cash on hand", "Time Deposits 3%", "Cert. of deposit, 5%(6MOS)",
      "Other current assets", "Accounts Rec. (net)", "Inventories",
      "Affiliate Receivable", "Subsidary Companies", "Furniture & Fixtures",
      "Equipment", "Building & Improvements", "Land",
      "Franchise Agreement", "Leased Property", "Accounts Payable",
      "Notes Payable, 13%", "Line of Credit, 15%", "Mortgage-Current portion",
      "Lease - Current portion", "Affiliate Payable", "Mortgage",
      "Capitalized Leases", "Common Stock @ 10 Par", "Additional Paid in Capital",
      "Retained Earnings/Deficit"
    ]
  
  # Dictionary to store extracted values
  extracted_values = {}
      
  # Loop through labels and find their values
  for label in labels:
      # Use regex to match the label and the number immediately after it
      match = re.search(rf"{label}\s+([\d.,]+)", page)  # Added ',' to capture numbers like 1,234.56
      # print(f"Searching for label: '{label}'")
      if match:
          extracted_values[label] = match.group(1)
          print(f"Found: {label} -> {match.group(1)}")
      else:
          extracted_values[label] = None  # Label not found
          print(f"Not found: {label}")
          
  # Verify the extracted dictionary
  print("Final Extracted Values:")
  print(json.dumps(extracted_values, indent=4))
  
  return extracted_values


# Main Program
# Creation of pdf for facilitator using the default output file name
text_to_pdf(input_file, output_file)


# Creation of pdf for each firm
for firm in firms:
  subset_file = path + "firm" + str(firm) + "_results.pdf"
  s = len(firms) * 2 + 1
  pages_to_keep = [firm * 2, firm * 2 + 1, s + firm]
  subset_pdf(output_file, subset_file, pages_to_keep)
  
  print(output_file)
  print(subset_file)
  print(pages_to_keep)


# Data extraction from P&L and Income Statement
# Read the file
with open(input_file, 'r') as file:
    content = file.read()

# Split into pages using '1' only when it's the first character in a line
pages = re.split(r'(?m)^1', content)

# Loop through the relevant pages to extract based on the number of firms
if len(pages) > 1:
  for firm in firms:
    print("Scanning firm " + str(firm))
    page = pages[firm * 2 - 1]          # page is a parameter for the page containint each firm's P&L
    
    # Extract values from the uploaded file
    extracted_pandl = extract_pl_is(page, "pl")
    extracted_incst = extract_pl_is(page, "is")
  
else:
  print("No pages to scan")

print("done")

