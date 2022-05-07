import argparse
import math
from pathlib import Path

from fillpdf import fillpdfs
from openpyxl import load_workbook
from fpdf import FPDF
from tqdm import tqdm


def fill_form():
    wb2 = load_workbook(filename='Vocabulario.xlsx')
    ws2 = wb2.active
    # save all the rows as a list of tupes
    rows = [(row[0].value, row[1].value) for row in ws2.iter_rows()]
    rows = rows[1:] # remove the first row header

    current_page = 0
    # open the pdf file and fill the fields with the values from the dictionary
    # get the number of fields in the pdf
    num_fields = fillpdfs.get_form_fields('3x3.pdf')
    #infer the dimensions of an array from the name of the fields, beginning in a_11 up to a_nn
    indices = [int(field[2:]) for field in num_fields]
    # get the maximum index
    max_index = max(indices)
    pdf_rows, pdf_cols = str(max_index)[0], str(max_index)[1]
    print(f"pdf is a {pdf_rows}x{pdf_cols} grid")
    numer_of_words_per_page = int(pdf_rows) * int(pdf_cols)
    fields = rows[current_page * numer_of_words_per_page:(current_page + 1) * numer_of_words_per_page]
    row_dicts = {f"a_{r}{c}": None for c in range(1, int(pdf_cols) + 1) for r in range(1, int(pdf_rows) + 1)}
    print(row_dicts)
    a_side = row_dicts.copy()
    b_side = row_dicts.copy()

    # fill the a_side dictionary with the first values of the rows
    for index, (key, _) in enumerate(fields):
        cell_index = f"a_{index // int(pdf_rows)+ 1}{index % int(pdf_cols) + 1}"
        a_side[cell_index] = key

    # fill the b_side dictionary with the second values of the rows
    for index, (_, value) in enumerate(fields):
        cell_index = f"a_{index // int(pdf_rows)+ 1}{int(pdf_cols)- index % int(pdf_cols) }"
        b_side[cell_index] = value
    print(a_side)
    print(b_side)
    fillpdfs.write_fillable_pdf('3x3.pdf', '3x3_aleman_1.pdf', a_side)
    fillpdfs.write_fillable_pdf('3x3.pdf', '3x3_spanisch_1.pdf', b_side)


def page_format_to_dimensions(page_format, orientation='P'):
    if page_format == 'A4':
        if orientation == 'P':
            return 297, 210
        elif orientation == 'L':
            return 210, 297


def create_pdf_page(pdf, text_values=[], num_rows=3, num_cols=3,
                    page_format='A4',
                    orientation='L', default_text=None):
    height, width = page_format_to_dimensions(page_format, orientation)
    #border_x = 10
    #border_y = 0
    pdf.add_page()
    rect_height = (height - pdf.get_y()*2) / num_rows
    rect_width = (width - pdf.get_x()*2) / num_cols
    #print(f"rect_height: {rect_height}")
    #print(f"rect_width: {rect_width}")
    pdf.set_font('Arial', 'B', 16)
    #print(f"x: {pdf.get_x()}, y: {pdf.get_y()}")
    index = 0
    for row in range(num_rows):
        #print("Incrementing Y")
        pdf.set_y(pdf.get_y())
        for col in range(num_cols):
            #print(f"x: {pdf.get_x()}, y: {pdf.get_y()}")
            line = 0 if col < num_cols - 1 else 1
            try:
                text = text_values[index]
            except IndexError:
                if default_text is None:
                    text = f"a_{row+1}{col+1}"
                else:
                    text = default_text
            pdf.cell(rect_width, rect_height, text, 1, line, 'C')
            index += 1


def extract_page_from_excel(excel_rows, page_number, page_rows, page_cols):
    page_size = page_rows * page_cols
    current_rows = excel_rows[page_number * page_size: (page_number + 1) * page_size]
    # pad the current rows with a tuple of two empty strings if they are not long enough
    current_rows += [('', '')] * (page_size - len(current_rows))
    #convert the first values in the rows to a list
    first_values = [row[0] for row in current_rows]
    #convert the second values in the rows to a list, inverting the rows
    second_values = []
    for i in range(page_rows):
        row_values = [row[1] for row in current_rows[i * page_cols: (i + 1) * page_cols]]
        row_values = row_values[::-1]
        second_values.extend(row_values)
    return first_values, second_values


def create_pdf_from_excel(excel_file, num_rows=3, num_cols=3, output_file='output.pdf', merge_sheets=False):
    if merge_sheets:
        pdf = FPDF(format='A4', orientation='L')
        pdf.set_auto_page_break(False)
    page_size = num_rows * num_cols
    wb2 = load_workbook(filename=excel_file, read_only=True)
    output_path = Path(output_file)
    for sheet in tqdm(wb2.worksheets):
        if not merge_sheets:
            pdf = FPDF(format='A4', orientation='L')
            pdf.set_auto_page_break(False)
        rows = [(row[0].value, row[1].value) for row in sheet.iter_rows()]
        rows = rows[1:]  # remove the first row header
        # the number of pages is the number of rows divided by the page size, rounded up
        num_pages = math.ceil(len(rows) / page_size)
        for page_number in range(num_pages):
            first_values, second_values = extract_page_from_excel(rows, page_number, num_rows, num_cols)
            create_pdf_page(pdf, text_values=first_values, default_text='', num_rows=num_rows, num_cols=num_cols)
            create_pdf_page(pdf, text_values=second_values, default_text='', num_rows=num_rows, num_cols=num_cols)
        if not merge_sheets:
            # append the name of the sheet to the output file and save it using pathlib
            sheet_output_file = output_path.with_name(f"{output_path.stem}_{sheet.title}.pdf")
            pdf.output(sheet_output_file)
    if merge_sheets:
        pdf.output(output_file, 'F')


if __name__ == '__main__':
    # read arguments from the command line and create the pdf
    parser = argparse.ArgumentParser(description='Create a pdf from an excel file')
    parser.add_argument('-f', '--file', type=str, help='The excel file to read from')
    parser.add_argument('-r', '--rows', type=int, help='The number of rows per page', default=3)
    parser.add_argument('-c', '--cols', type=int, help='The number of columns per page', default=3)
    parser.add_argument('-o', '--output', type=str, help='The output file')
    parser.add_argument('-m', '--merge', action='store_true', help='Merge all sheets into one pdf', default=False)
    args = parser.parse_args()
    create_pdf_from_excel(args.file, args.rows, args.cols, args.output, merge_sheets=args.merge)
    print("Done")
