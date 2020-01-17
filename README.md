# nrb-costing-sheets
## To Edit Source Code and Work with GIT
1. Use Git Bash
2. `cd ../../Development`
3. `git clone https://github.com/northriverboats/nrb-costing-sheets.git`
4. `cd nrb-costing-sheets`

5. Use windows shell
6. `cd \Development\nrb-costing-sheets`
7. `\Python37\python -m venv .venv`
8. `.venv\Scripts\activate`
9. `python -m pip install pip --upgrade`
10. `pip install -r requirements.txt`
11. Remember to Create New Branch Before Doing Any Work

## Build Executable
`.venv\Scripts\pyinstaller.exe --onefile --windowed --icon options.ico  --name "Excel Costing Sheets" "Excel Costing Sheets FWW.spec" main.py`
