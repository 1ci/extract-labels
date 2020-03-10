#!/usr/bin/env python
import sys
import io
import os
import re
import pdfminer.high_level
import pdfminer.layout
import openpyxl
from openpyxl.styles import Font, Color, Alignment, Border, Side, colors

FILEPATH = os.path.dirname(os.path.realpath(__file__)) + '/' + './file.pdf'
SEARCH_PATTERN = r"(?:[-,.\/\\а-яА-Я0-9_\t ]+\n?)+"
LABELS_PER_COL = 5
COLUMN_WIDTH = 19.4
ROW_HEIGHT = 60
FONT_SIZE = 8
FONT_NAME = "Arial" # Arial, Times New Roman

def export_labels(pdf: str) -> list:
    print("Attempting to parse pdf file:", pdf)
    with open(pdf, 'rb') as f:
        output_string = io.StringIO()
        laparams = pdfminer.layout.LAParams() # default layout parameters
        try:
            pdfminer.high_level.extract_text_to_fp(f, output_string, laparams=laparams,
                page_numbers=list(range(1, 30)))

            print("Parsed pdf file successfully.")
        except: # usually pdfminer.pdfparser.PDFSyntaxError
            print("Could not parse pdf file:", pdf)
        
        print("Extracting labels...")
        string = output_string.getvalue().strip()
        pattern = re.compile(SEARCH_PATTERN)
        matches = re.finditer(pattern, string)
        labels = []
        for match in matches:
            print(match)
            labels.append(match.group(0).strip())
        return labels

def create_spreadsheet(labels: list, filename: str):
    print("Creating spreadsheet...")
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # a bit of styling
    center_aligned_text = Alignment(horizontal="center",
                                     vertical='center',
                                     wrap_text=True)

    for col in ['A', 'B', 'C', 'D', 'E']:
        sheet.column_dimensions[col].width = COLUMN_WIDTH
    
    column, row = 0, 0
    for i, label in enumerate(labels):
        if (column := i % LABELS_PER_COL) == 0:
            row = row + 1
            sheet.row_dimensions[row].height = ROW_HEIGHT

        sheet.cell(row, column + 1, label).alignment = center_aligned_text
        sheet.cell(row, column + 1).font = Font(size = FONT_SIZE, name = FONT_NAME) 
    
    # fix margins
    sheet.page_margins.left = 0.39
    sheet.page_margins.right = 0.39
    sheet.page_margins.top = 0.39
    sheet.page_margins.bottom = 0.39
    sheet.page_margins.header = 0
    sheet.page_margins.footer = 0
    sheet.print_options.horizontalCentered = True
    
    try:
        filename = filename.replace(".pdf", ".xlsx")
        workbook.save(filename = filename)
        print("Saved spreadsheet:", filename)
    except PermissionError:
        print("Could not open file for writing:", filename)
        print("Please make sure the file is not being read by another program.")
    except:
        print("Failed to save spreadsheet:", filename)

def main(args: list):
    if len(args) < 2:
        print('Usage: ./extract-labels <pdffile1> <pdffile2> ...')
        input("Press ENTER to exit.")
        return
    
    for i, a in enumerate(args):
        if i == 0:
            continue
        labels = export_labels(a)
        print("Number of labels:", len(labels))
        #print(labels)
        create_spreadsheet(labels, a)
    
    input("Job finished.")

if __name__ == '__main__':
    main(sys.argv)
