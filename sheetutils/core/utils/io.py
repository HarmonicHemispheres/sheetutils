import openpyxl
import xlrd3 as xlrd
from pathlib import Path
from typing import Union



def open_xls(file: Union[Path,str]) -> xlrd.Book:

    file = Path(file)

    if file.exists(): 
        return xlrd.open_workbook(file)


def open_xlsx(file: Union[Path,str]) -> openpyxl.Workbook:

    file = Path(file)

    if file.exists(): 
        return openpyxl.open(file)

