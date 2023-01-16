
from sheetutils.core.cell import Cell
from sheetutils.core.range import Range
import sheetutils.core.utils.sheetops as shops
from typing import Union, List, Iterator
from pathlib import Path
import xlrd3 as xlrd
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
import openpyxl.cell.cell as openpyxl_cell
from abc import ABC, abstractmethod, abstractproperty


##########################################
### BASE CLASS
##########################################
class Sheet(ABC):
    
    @abstractproperty
    def name(self) -> str:
        pass

    @abstractproperty
    def workbook(self) -> Path:
        pass
    
    @abstractmethod
    def cell(self, address: str=None, row:int=None, col:int=None) -> Cell:
        pass
    
    @abstractmethod
    def cells(self, address: str) -> Iterator[Cell]:
        pass
    
    @abstractmethod
    def range(self, address: str) -> Range:
        pass

        

##########################################
### XLSX (SHEET) CLASS
##########################################
class Xlsx(Sheet):
    def __init__(self, data: Worksheet=None, workbook: Path = None):
        self._data: Worksheet = data
        self._name = self._data.title
        self._workbook = Path(workbook) if workbook else workbook

    @property
    def name(self) -> str:
        return self._name

    @property
    def workbook(self) -> Path:
        return self._workbook

    def cell(self, address: str=None, row:int=None, col:int=None) -> Cell:
        # -- Validation
        if address:
            col,row = shops._addr_to_tuple(address)
        elif not row or not col:
            raise Exception("if no address is used, row and col must be provided!")
        
        # -- get cell
        cell = self._data.cell(row, col)

        # -- return result
        return Cell(val=cell.value, row=row, col=col)
    
    def cells(self) -> Iterator[Cell]:
        for row in self._data:
            for cell in row:
                yield Cell(val=cell.value, row=cell.row, col=cell.column)

    def range_all(self) -> Range:
        rng = Range()
        rng._sheet = self.name
        rng._workbook = self.workbook
        cells = tuple(self.cells())
        rng._a1_cell = cells[0]
        rng._a2_cell = cells[-1]
        for cell in cells:
            rng._add_cell_to_cache(cell)

        return rng

    
    def range(self, address: str) -> Range:
        rng = Range(address=address)
        rng._sheet = self.name
        rng._workbook = self.workbook

        for r,c in rng._iter_cell_rcs():
            cell = self.cell(row=r, col=c)
            rng._add_cell_to_cache(cell)
        
        return rng
    
    def get_after_in_row(self, address:str, key:str):
        """returns the first non-null value after a key in a row"""
        pass
        



##########################################
### XLS (SHEET) CLASS
##########################################
class Xls(Sheet):
    def __init__(self, data: xlrd.sheet.Sheet=None, workbook: Path = None):
        self._data: xlrd.sheet.Sheet = data
        self._name = self._data.name
        self._workbook = Path(workbook) if workbook else workbook

    @property
    def name(self) -> str:
        return self._name

    @property
    def workbook(self) -> Path:
        return self._workbook
        
    def cell(self, address: str=None, row:int=None, col:int=None) -> Cell:
        # -- Validation
        if address:
            col,row = shops._addr_to_tuple(address)
        elif not row or not col:
            raise Exception("if no address is used, row and col must be provided!")
        
        # -- get cell
        value = self._data.cell_value(rowx=row-1, colx=col-1)

        # -- return result
        return Cell(val=value, row=row, col=col)
    
        
    def cells(self) -> Iterator[Cell]:
        for r in range(0, self._data.nrows):
            for c in range(0, self._data.ncols):                
                val = self._data.cell_value(r, c) or None
                yield Cell(val=val, row=r+1, col=c+1)
    
    def range(self, address: str) -> Range:
        rng = Range(address=address)
        rng._sheet = self.name
        rng._workbook = self.workbook

        return rng


##########################################
### SHEET COLLECTION CLASS
##########################################
class SheetCollection:
    def __init__(self):
        self._sheets = {}


    def get(self, name: str):
        sheet = self._sheets.get(name)
        return sheet
        
    def add(self, sheet: Union[Xls,Xlsx]):
        self._sheets[sheet.name] = sheet