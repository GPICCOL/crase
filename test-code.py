# Clear all variables from the current Python namespace
globals().clear()

import os
import re
import json
import pandas as pd
from openpyxl import load_workbook

path = "./input-ready/"
file = "input-mkt-firm1.xlsx"
workbook = load_workbook(path + file, data_only=True)
sheet = workbook["Marketing Plan"]

print(sheet["E10"].value)
infoguest = 1 if sheet["E42"].value == "Yes" else 0
infofin = 1 if sheet["G42"].value == "Yes" else 0
infoprod = 1 if sheet["J42"].value == "Yes" else 0

menu_rows = [10 + 2 * i for i in range(8)]
menu_items = [sheet[f"A{row}"].value for row in menu_rows]
menu_preps = [sheet[f"E{row}"].value for row in menu_rows]
pnps = [sheet[f"C{row}"].value for row in menu_rows]
portions = [sheet[f"G{row}"].value for row in menu_rows]
mkt = [sheet[f"J{row}"].value for row in menu_rows]
prices = [sheet[f"L{row}"].value for row in menu_rows]

print(menu_items)
print(menu_preps)

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

menu = menu.astype({
    "menu_item": "string",      # Enforce as string type
    "menu_prep": "string",      # Enforce as string type
    "pnp": "int",             # Nullable integer type
    "portion": "float",         # Float type
    "marketing": "int",       # Float type
    "price": "float",           # Float type
    "food_cost": "float",       # Float type
})

f_name = "test.txt"
with open(f_name, "w") as file:
    file.write("\r\n\r\n\r\n")
  
menu_item_line = []
menu_item_line_count = 0
for _, row in menu.iterrows():
    # Format the string with exact positions
    menu_item_line = (
        f"{row['menu_item']:<32}"  # menu_item starts at the beginning
        f"{row['menu_prep']:<20}"  # menu_prep starts at character 32
        f"{row['pnp']:>1}"         # pnp starts at character 52
        f"{row['portion']:>10.1f}" # portion starts at character 58
        f"{row['marketing']:>10}"  # marketing starts at character 68
        f"{row['price']:>10.2f}"    # price starts at character 79
        f"{row['food_cost']:>10.3f}"    # price starts at character 88
    )
    with open(f_name, "a") as file:  # Open the file in append mode
      file.write(menu_item_line + "\r\n")  # Add the string with a newline
      menu_item_line_count += 1  # Increment the line count

