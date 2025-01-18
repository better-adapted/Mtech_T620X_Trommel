# CODESYS PYTHON TOOLS


## Python

run `python` in a powershell prompt. If the marketplace pops up then get it installed.

## venv

### Create

1. change directory to this project folder
2. Run `python -m venv ./.venv` to create enviroment

### Activate enviroment

running `./.venv/Scripts/activate` will cause `(.venv)` to appear before your current path indicating you are in the python virtual enviroment

### Install dependencies

When you are in the virtual enviroment then install dependencies via `pip` with `pip install -r ./requirement.txt`

## Scripts

### `read_excel.py`

- Make sure your are in the `Codesys_Python_Tools` directory

`python read_excel.py`