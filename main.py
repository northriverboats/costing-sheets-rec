import click
import sys
from pickle import load
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

dbg = 0


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

def unpickle_boats(folder):
    if folder == None:
        folder = r"K:\Links\2020\Boats"
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
    boats = unpickle_boats(folder)
    print(len(boats))


if __name__ == "__main__":
    main()
