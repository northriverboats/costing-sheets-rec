# nrb-costing-sheets
## To Edit Source Code and Work with GIT
1. Use Git Bash
2. `cd ../../Development`
3. `git clone git@github.com:northriverboats/costing-sheets-rec.git`
4. `cd costing-sheets-rec`

5. Use windows shell
6. `cd \Development\costing-sheets-rec`
7. `\Python37\python -m venv .venv`
8. `.venv\Scripts\activate`
9. `python -m pip install pip --upgrade`
10. `pip install -r requirements.txt`
11. Remember to Create New Branch Before Doing Any Work

## Build Executable
`.venv\Scripts\pyinstaller.exe --onefile --windowed --icon options.ico  --name "Costing Sheets for Rec" "costing-sheet-rec.spec" main.py`