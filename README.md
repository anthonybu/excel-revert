# excel-revert
Download pre-built executable in [releases](https://github.com/anthonybu/excel-revert/releases)

**TL;DR** Use this tool to revert sort excel content without manually adding index column. Support convert single file or batch a given folder. Output will be saved in the "reverted" subdirectory with the same file name.

# background story
An accountant friend needs to reconcile local ledger with custodian bank records on daily basis. New transaction is always appended to the end of the ledger, but bank records always have latest transaction on top. This requires a manual pre-process of adding index column to the bank records, revert sort using the index, then remove the column.

# build executables
Build with **PyInstaller** `python -OO -m PyInstaller --clean --onefile --icon=excel_revert.ico revert.py`
