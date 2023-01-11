
from sheetutils.core.cell import Cell
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
    def name(self, address: str) -> str:
        pass
    
    @abstractmethod
    def cell(self, address: str) -> Cell:
        pass
    
    @abstractmethod
    def cells(self, address: str) -> Iterator[Cell]:
        pass


##########################################
### XLSX (SHEET) CLASS
##########################################
class Xlsx(Sheet):
    def __init__(self, data: Worksheet=None):
        self._data: Worksheet = data

        self._name = self._data.title

    @property
    def name(self) -> str:
        return self._name

    def cell(self, address: str) -> Cell:
        row, col = shops._addr_to_tuple(address)
        cell = self._data.cell(row, col)
        return Cell(val=cell.value, row=row, col=col)
    
    def cells(self) -> Iterator[Cell]:
        for row in self._data:
            for cell in row:
                yield Cell(val=cell.value, row=cell.row-1, col=cell.column-1)


##########################################
### XLS (SHEET) CLASS
##########################################
class Xls(Sheet):
    def __init__(self, data: xlrd.sheet.Sheet=None):
        self._data: xlrd.sheet.Sheet = data
        self._name = self._data.name

    @property
    def name(self) -> str:
        return self._name
        
    def cell(self, address: str) -> Cell:
        row, col = shops._addr_to_tuple(address)
        value = self._data.cell_value(rowx=row-1, colx=col-1)
        return Cell(val=value, row=row, col=col)
    
        
    def cells(self) -> Iterator[Cell]:
        for row_idx in range(0, self._data.nrows):
            for col_idx in range(0, self._data.ncols):
                val = self._data.cell_value(row_idx, col_idx) or None
                yield Cell(val=val, row=row_idx, col=col_idx)


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