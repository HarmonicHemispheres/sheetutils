
from pydantic import BaseModel
from sheetutils.core.cell import Cell
from typing import List, TYPE_CHECKING
from copy import deepcopy
from datetime import datetime
import time

class Table:
    def __init__(self, model: BaseModel = None):
        self._model = model

        self.entries: List[BaseModel] = []
        
        #  <column> : {<key> : <index in self.entries>}
        self.indexes = {
            "_base_column": {
                "_key": 0
            }
        } 

    # ------ PRIVATE METHODS
    def _filter(self, query:str): pass
    def _get_column(self): pass
    def _get_column_values(self): pass

    # ------ PROPERTIES METHODS
    ...

    # ------ PUBLIC METHODS
    def from_range(self, 
                   rng: "Range", 
                   header_row: int=0, 
                   ) -> "Table":
        new_table = deepcopy(self)

        # parse header
        col_indexes = {}
        for row in rng.iter_rows(start=header_row):
            for cidx,col in enumerate(row):
                col_indexes[col.val] = cidx
            break

        # parse data
        for row in rng.iter_rows(start=header_row+1):
            entry_dict = {k:row[idx].val for k,idx in col_indexes.items()}
            new_table.add(data=entry_dict)

        return new_table

    def add(self, data: dict):
        self.entries.append(
            self._model.parse_obj(data)
        )
        