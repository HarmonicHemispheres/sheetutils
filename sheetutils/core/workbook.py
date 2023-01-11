
import sheetutils.core.utils.io as io
from sheetutils.core.sheet import Sheet, Xlsx, Xls, SheetCollection
from typing import Union
from pathlib import Path
import xlrd3 as xlrd
import openpyxl
from enum import Enum

class WorkbookType(Enum):
    unknown = 0
    xlsx = 1
    xls = 2

class Workbook:
    def __init__(self, 
                 file: Union[str,Path]=None,
                 *args,
                 **kwargs
                 ):
        # public vars
        self.file: Path = Path(file)
        self.type: WorkbookType = WorkbookType.unknown

        # private vars
        self._sheets: SheetCollection = SheetCollection()

        self._raw_xlsx: openpyxl.Workbook = None
        self._raw_xls: xlrd.Book = None


    def load(self):
        if self.file.suffix == ".xlsx":
            self.type = WorkbookType.xlsx
            self._raw_xlsx = io.open_xlsx(self.file)
        elif self.file.suffix == ".xls":
            self.type = WorkbookType.xls
            self._raw_xls = io.open_xls(self.file)
        
        return self

    def load_sheet(self, name: str) -> Sheet:
        if self.type == WorkbookType.xlsx:
            return Xlsx(self._raw_xlsx.get_sheet_by_name(name))
        elif self.type == WorkbookType.xls:
            for idx, sheet_name in enumerate(self._raw_xls.sheet_names()):
                if sheet_name == name:
                    return Xls(self._raw_xls.sheet_by_index(idx))
        

    def sheet(self, name: str) -> Sheet:
        # check for cached sheet
        sheet = self._sheets.get(name)
        if sheet:
            return sheet
        
        # otherwise attempt to load 
        sheet_loaded = self.load_sheet(name)
        self._sheets.add(sheet_loaded)

        return sheet_loaded


