# Clear all variables from the current Python namespace
globals().clear()

import os
import re
import json
import pandas as pd
from openpyxl import load_workbook

# Read firm marketing plan
path = "./input-ready/"
files = os.listdir(path)
file = files[0]

def make_filename(file):
  # Load an Excel file
  workbook = load_workbook(path + file)
  sheet = workbook["Marketing Plan"]
  
  # Read firm, market area, season data
  firm_number = str(re.search(r'\d', file).group())
  market_area = sheet["C44"].value
  restaurant_name = sheet["J2"].value
  season = sheet["J4"].value
  
  # Create needed variables
  f_name = "RFI" + firm_number + market_area + ".TXT"  # Concatenate the parts to form the file name
  
  # Open the file in write mode and add three empty lines
  with open(f_name, "w") as file:
      file.write("\n\n\n")
  
  print(f"File '{f_name}' has been created with three empty lines.")
  
  # Create the line containing the restaurant name with appropriate spacing
  total_length = 84
  restaurant_string = restaurant_name.center(total_length) + "LS        B1"
  
  # Append the restaurant name to the file
  with open(f_name, "a") as file:  # Open the file in append mode
      file.write(restaurant_string + "\n")  # Add the string with a newline
  
  print(f"'{restaurant_string}'")  # Prints the result with quotes to visualize spaces
  
  return f_name

def make_menu(file, f_name):
  # Load an Excel file
  workbook = load_workbook(path + file)
  sheet = workbook["Marketing Plan"]
  
  # Read menu data
  menu_rows = [10 + 2 * i for i in range(8)]
  menu_items = [sheet[f"A{row}"].value for row in menu_rows]
  menu_preps = [sheet[f"C{row}"].value for row in menu_rows]
  pnps = [2, 2, 2, 2, 2, 2, None, None]
  portions = [sheet[f"E{row}"].value for row in menu_rows]
  mkt = [sheet[f"H{row}"].value for row in menu_rows]
  prices = [sheet[f"J{row}"].value for row in menu_rows]

  # Menu saved as a pandas dataframe
  menu = pd.DataFrame({
      "menu_item": menu_items,
      "menu_prep": menu_preps,
      "pnp": pnps,
      "portion": portions,
      "marketing": mkt,
      "price": prices,
      "food_cost": [0.000, 0.000, 0.000, 0.000, 0.000, 0.000, None, None],
  })
  menu = menu[menu["menu_item"].notna()]
  
  # Append menu decisions
  # Create each line and append it
  menu_item_line = []
  menu_item_line_count = 0
  for _, row in menu.iterrows():
      # Format the string with exact positions
      menu_item_line = (
          f"{row['menu_item']:<32}"  # menu_item starts at the beginning
          f"{row['menu_prep']:<20}"  # menu_prep starts at character 32
          f"{row['pnp']:<6}"         # pnp starts at character 52
          f"{row['portion']:<11.1f}" # portion starts at character 58
          f"{row['marketing']:<10}"  # marketing starts at character 68
          f"{row['price']:<9.1f}"    # price starts at character 79
          f"{row['food_cost']:<6.3f}"    # price starts at character 88
      )
      with open(f_name, "a") as file:  # Open the file in append mode
        file.write(menu_item_line + "\n")  # Add the string with a newline
        menu_item_line_count += 1  # Increment the line count
  
  # Add the required number of empty lines to get to 10 menu items
  with open(f_name, "a") as file:
      for _ in range(10 - menu_item_line_count):  # Loop for the required number of lines
          padded_line = f"{'':<88}0.000\n"  # Pad with spaces to 87 characters, then add "0.000"
          file.write(padded_line)
          
  print(f"File '{f_name}' now contains 10 lines, with {menu_item_line_count} menu items and the remaining as blank lines.")
  
  return menu

def make_ops(file, f_name):
  # Define the additional lines to be added
  additional_lines = [
      "Management                                        1.5       0.0       0.0         0         0         0         0         0         0         0",
      "Supervisory                                       1.0       0.0       0.0         0         0         0         0         0         0         0",
      "Service                                          10.0       0.0       0.0         0         0",
      "Support                                           9.0       0.0       0.0         0         0",
      "Operations                                       36.0     10.25      0.28      6500     19500         0         0         0         0     12000",
      "Finance                                             0         0         0         0         0         0         0         0         0         0",
      "Assets - Investments                                0         0",
      "       - Leases                                     0         0",
      "       - Borrowings                                 0         0",
      "Variances",
  ]
  
  # Write the additional lines to the file
  with open(f_name, "a") as file:
      for line in additional_lines:
          file.write(line + "\n")

for file in files:
  print(file)
  f_name = make_filename(file)
  menu = make_menu(file, f_name)
  make_ops(file, f_name)




