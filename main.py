import click
import sys
import os
from pickle import load
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from dotenv import load_dotenv

dbg = 0
consumable_column = 9

rates = [
    # labor rate, column, row
    ["FABRICATION LABOR RATE", 7, 428],
    ["PAINT LABOR RATE", 7, 429],
    ["OUTFITTING LABOR RATE", 7, 431],
]
sections = [
    # Section, start, end, consumable, start-del, end-del
    ["FABRICATION", 8, 32, 35, 0, 0],
    ["PAINT", 42, 64, 67, 0, 0],
    ["CANVAS", 74, 96, 0, 71, 102],
    ["OUTFITTING", 106, 356, 0, 0, 0 ],
    ["ENGINE & JET", 391, 401, 0, 388, 407],
    ["TRAILER", 411, 412, 0, 0, 0],
]


"""
Levels
0 = no output
1 = minimal output
2 = verbose outupt
3 = very verbose outupt
"""
def debug(level, text):
    if dbg > (level -1):
        print(text)

def load_environment():
    if getattr(sys, 'frozen', False):
        bundle_dir = sys._MEIPASS # pylint: disable=no-member
    else:
        # we are running in a normal Python environment
        bundle_dir = os.path.dirname(os.path.abspath(__file__))

    # load environmental variables
    load_dotenv(dotenv_path = Path(bundle_dir) / ".env")

def unpickle_boats(pickle_folder):
    base = Path(pickle_folder)

    pickle_file = base.joinpath(base.stem + '.pickle')
    if pickle_file.exists() == False:
        debug(0, str(pickle_file) + " not found.")
        sys.exit(1)

    boats = load(open(pickle_file, 'rb'))
    return boats

def resolve_environment(value, environmental_string):
    if value == None:
        value = os.getenv(environmental_string)
    return value


def process_labor_rate(ws, boats, model):
    for rate, column, row in rates:
        labor = float(boats[model][rate])
        _ = ws.cell(column=column, row=row, value=labor)

def process_part_highlighting(ws, length, part):
    mode = part[str(length) + ' RRS']
    if "P" in mode:
        pass
    if "Z" in mode:
        pass

def process_by_parts(ws, boats, model, length, section, start, end):
    for part in sorted(boats[model][section[0] + ' PARTS'], key = lambda i: (i['VENDOR'], i['PART NUMBER'])):
        qty = float(part[str(length) + ' QTY'])
        if qty > 0: 
            print('           | {:15.15} | {:15.15} | {:25.25} | {:8.2f} | {:6.6} | {:10.4f} | {:14.6f} | {:2.2} |'.format(
                part['VENDOR'],
                part['PART NUMBER'][1:-1],
                part['DESCRIPTION'],
                part['PRICE'],
                part['UOM'],
                float(part[str(length) + ' QTY']),
                float(part[str(length) + ' TOTAL']),
                (part[str(length) + ' RRS']),
            ))
            process_part_highlighting(ws, length, part)

def delete_unused_section(ws, start_del, end_del):
    ws.delete_rows(start_del, end_del - start_del)

def process_consumables(ws, boats, model, length, section, start_row, consumable_row):
    if consumable_row > 0:
        consumables = float(boats[model][section + ' CONSUMABLES'])
        formula = "=SUM({}{}:{}{})*{}".format(
            chr(64 + consumable_column),
            start_row,
            chr(64 + consumable_column),
            consumable_row - 1,
            consumables
        )
        _ = ws.cell(column=consumable_column, row=consumable_row, value=formula)
    

def process_by_section(ws, boats, model, length):
    for section, start_row, end_row, consumable_row, start_delete_row, end_delete_row in sections[::-1]:
        number_of_parts = len(boats[model][section + ' PARTS'])
        if number_of_parts == 0 and start_delete_row > 0:
            delete_unused_section(ws, start_delete_row, end_delete_row)
        else:
            process_consumables(ws, boats, model, length, section, start_row, consumable_row)
            # process_by_parts(ws, boats, model, length, section, start_row, end_row)


def generate_filename(model, length, output_folder):
    return output_folder + "\\" + "Costing Sheet {}' {}.xlsx".format(length, model.upper())

def process_sheetname(ws, model, length):
     _ = ws.cell(column=3, row=3, value="{}' {}".format(length, model))

def load_template(template_file):
    wb = load_workbook(template_file)
    ws = wb.active
    return [wb, ws]

def process_boat(boats, model, length, output_folder, template_file):
    wb, ws = load_template(template_file)
    
    process_sheetname(ws, model, length)
    process_labor_rate(ws, boats, model)
    process_by_section(ws, boats, model, length)

    file_name = generate_filename(model, length, output_folder)
    wb.save(file_name)

def process_by_length(boats, model, output_folder, template_file):
    for length in boats[model]["BOAT SIZES"]:
        process_boat(boats, model, length, output_folder, template_file)
        break

def process_by_model(boats, output_folder, template_file):
    for model in boats:
        process_by_length(boats, model, output_folder, template_file)
        break

# pylint: disable=no-value-for-parameter
@click.command()
@click.option('--debug', '-d', is_flag=True, help='show debug output')
@click.option('--verbose', '-v', default=1, type=int, help='verbosity level 0-3')
@click.option('--folder', '-f', required=False, type=click.Path(exists=True, file_okay=False), help="directory to process")
@click.option('--output', '-o', required=False, type=click.Path(exists=True, file_okay=False), help="output directory")
@click.option('--template', '-t', required=False, type=click.Path(exists=True, dir_okay=False), help="template xlsx sheet")
def main(debug, verbose, folder, output, template):
    load_environment()
    pickle_folder = resolve_environment(folder, 'FOLDER')
    output_folder = resolve_environment(output, 'OUTPUT')
    template_file = resolve_environment(template, 'TEMPLATE')

    boats = unpickle_boats(pickle_folder)

    process_by_model(boats, output_folder, template_file)


if __name__ == "__main__":
    main()