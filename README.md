# nrb-costing-sheets
## To Edit Source Code and Work with GIT
1. Use Git Bash
2. `cd ../../Development`
2. `git clone https://github.com/northriverboats/nrb-costing-sheets.git`
2. `cd nrb-costing_sheets`
2. Use windows shell
2. `cd \Development\nrb-costing_sheets`
3. `\Python37\python -m venv .venv`
4. `.venv\Scripts\activate`
5. `python -m pip install pip --upgrade`
6. `pip install -r requirements.txt`
7. patch pyinstall's bindepend.py as noted below.
8. Remember to Create New Branch Before Doing Any Work

## Build Executable
`.venv\Scripts\pyinstaller.exe --onefile --windowed --icon options.ico  --name "Excel Costing Sheets" "Excel Costing Sheets FWW.spec" main.py`


## Pyinstall Patch
`vim lib\site-packages\PyInstaller\depend\bindepend.py`  
after line 874  
`# Python library NOT found. Resume searching using alternative methods.`
```
    # Work around for python venv having VERSION.dll rather than pythonXY.dll
    if is_win and 'VERSION.dll' in dlls:
        pydll = 'python%d%d.dll' % sys.version_info[:2]
        if pydll in PYDYLIB_NAMES:
            filename = getfullnameof(pydll)
            return filename
```
