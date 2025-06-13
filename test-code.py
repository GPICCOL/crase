import os
import re
import json
import pandas as pd
from openpyxl import load_workbook

# Read firm marketing plan
path = "./input-ready/"
file = "input-master-all-firms.xlsx"

def make_fin(file):
  # Load an Excel file
  workbook = load_workbook(path + file, data_only=True)
  sheet = workbook["Financial Plan"]
  
  # Read firm, market area, season data
  timedep = str(sheet["D7"].value) + str(sheet["E7"].value)
  certdep = str(sheet["D9"].value) + str(sheet["E9"].value)
  notepay = str(sheet["D11"].value) + str(sheet["E11"].value)
  capital = str(sheet["D13"].value) + str(sheet["E13"].value)
  dividend = str(sheet["D15"].value) + str(sheet["E15"].value)
  
  # Create financial decisions line
  fin_decisions = (
    f"{'Finance':<43}"   # Financial left aligned, padded to 43 characters
    f"{timedep:>10}"         # sales are right aligned to position 53
    f"{certdep:>10}"    # food cost is 10 characters right aligned 2 decimal
    f"{notepay:>10}"         # beverage cost is 10 characters right aligned 2 decimal
    f"{capital:>10}"          # labor cost is 10 characters right aligned integer
    f"{dividend:>10}" # other costs is 10 characters right aligned integer
    f"{'0':>10}"           # unknown padding 0 value is 10 characters right aligned
    f"{'0':>10}"           # unknown padding 0 value is 10 characters right aligned
    f"{'0':>10}"           # unknown padding 0 value is 10 characters right aligned
    )
  return fin_decisions


print(make_fin(file))




