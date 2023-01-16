from pydantic import BaseModel
from sheetutils.core.cell import Cell
from sheetutils.core.utils.sheetops import _addr_to_tuple
from typing import Tuple, Dict, Iterator
from pathlib import Path


class Range:
    def __init__(self, address:str=None):
        self._a1_cell: Cell = None
        self._a2_cell: Cell = None
        self._cells: Dict[Tuple[int,int],Cell] = {}    # Row, Col
        self._address: str = address
        self._sheet: str = None
        self._workbook: Path = None

        if self._address:
            self._a1_cell, self._a2_cell = self._cells_from_address(self._address)

    def __repr__(self) -> str:
        return f"<RANGE: A:{self._address} SH:{self._sheet} WB:{self._workbook} VCELLS:{self.vcount} CELLS:{self.count}>"

    # ------ PRIVATE METHODS
    def __iter__(self):
        for cell in self.iter_cells():
            yield cell

    def _add_cell_to_cache(self, cell:Cell):
        self._cells[(cell.row, cell.col)] = cell

    def _cells_from_address(self, address:str) -> Tuple[Cell, Cell]:
        a1,a2 = self._address.replace("$","").split(":")
        cr1 = _addr_to_tuple(a1)
        cr2 = _addr_to_tuple(a2)

        return (
            Cell(row=cr1[1], col=cr1[0]),
            Cell(row=cr2[1], col=cr2[0])
        )
    
    def _iter_cell_rcs(self) -> Iterator[Tuple[int,int]]:
        for c in range(self._a1_cell.col, self._a2_cell.col+1):
            for r in range(self._a1_cell.row, self._a2_cell.row+1):
                yield (r,c)


    # ------ PROPERTIES METHODS
    @property
    def count(self) -> int:
        """returns the cell count. cells based on values extracted and cached"""
        return len(self._cells)
    
    @property
    def vcount(self) -> int:
        """returns the virtual cell count. cells based on the range address"""
        # print("A1: ", self._a1_cell)
        # print("A2: ", self._a2_cell)
        # print("C: ", (self._a2_cell.col+1-self._a1_cell.col))
        # print("R: ", (self._a2_cell.row+1-self._a1_cell.row))
        return (self._a2_cell.col+1-self._a1_cell.col) * (self._a2_cell.row+1-self._a1_cell.row)

    # ------ PUBLIC METHODS
    def iter_cells(self) -> Iterator[Cell]:
        for r in range(self._a1_cell.row, self._a2_cell.row+1):
            for c in range(self._a1_cell.col, self._a2_cell.col+1):
                yield self._cells[(r,c)]

    def iter_cols(self, start: int=0) -> Iterator[Cell]:
        for c in range(self._a1_cell.col+start, self._a2_cell.col+1):
            yield [self._cells[(r,c)] for r in range(self._a1_cell.row, self._a2_cell.row+1)]
                

    def iter_rows(self, start: int=0) -> Iterator[Cell]:
        for r in range(self._a1_cell.row+start, self._a2_cell.row+1):
            yield [self._cells[(r,c)] for c in range(self._a1_cell.col, self._a2_cell.col+1)]

    def find_kv(self, key:str, 
                      val_direction="right", 
                      ignore_empty_cells=False
                      ) -> Cell:
        """..."""
        for cell in self.iter_cells():
            if cell.val == key:
                while cell:
                    if val_direction == "right":
                        val_cell = self._cells.get((cell.row, cell.col+1))
                    elif val_direction == "left":
                        val_cell = self._cells.get((cell.row, cell.col-1))
                    elif val_direction == "down":
                        val_cell = self._cells.get((cell.row+1, cell.col))
                    elif val_direction == "up":
                        val_cell = self._cells.get((cell.row-1, cell.col))
                    
                    if ignore_empty_cells:
                        if val_cell and val_cell.val in ("",None):
                            cell = val_cell
                        else:
                            return val_cell
                    else:
                        return val_cell

        return None

        
