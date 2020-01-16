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
    ["CANVAS", 74, 96, 0, 0, 0],
    ["OUTFITTING", 106, 356, 0, 0,0 ],
    ["ENGINE & JET", 391, 401, 0, 388, 407],
    ["TRAILER", 411, 412, 0, 408, 413],
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

def unpickle_boats(folder):
    if folder == None:
        folder = os.getenv('FOLDER')
    base = Path(folder)

    pickle_file = base.joinpath(base.stem + '.pickle')
    if pickle_file.exists() == False:
        debug(0, str(pickle_file) + " not found.")
        sys.exit(1)

    boats = load(open(pickle_file, 'rb'))
    return boats

def find_excel_files(folder):
    pass

# pylint: disable=no-value-for-parameter
@click.command()
@click.option('--debug', '-d', is_flag=True, help='show debug output')
@click.option('--verbose', '-v', default=1, type=int, help='verbosity level 0-3')
@click.option('--folder', '-f', required=False, type=click.Path(exists=True, file_okay=False), help="folder to process")
def main(debug, verbose, folder):
    load_environment()
    boats = unpickle_boats(folder)
    print(len(boats))

print('this works!!!!!!!!')
if __name__ == "__main__":
    print('this works!!!!')
    main()
