# Clear all variables from the current Python namespace
globals().clear()

import os
import re
import json
from fpdf import FPDF
from PyPDF2 import PdfReader, PdfWriter

# Initialization parameters
path = "./output-dashboard/"
input_file = path + "FR.LST"
output_file = path + "facilitator.pdf"
firms = [1, 2, 3]



