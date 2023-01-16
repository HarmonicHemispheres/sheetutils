
from pydantic import BaseModel, create_model
from sheetutils.core.cell import Cell
from typing import List, TYPE_CHECKING
from copy import deepcopy
from datetime import datetime
from pathlib import Path
import pandas as pd
import time
from uuid import uuid4

class Table:
    def __init__(self, model: BaseModel = None):
        self._model = model
        self.entries: List[BaseModel] = []
        

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
                   dynamic_col_headers=True,
                   skip_nulls: bool = False,
                   include_sys_params: bool = False
                   ) -> "Table":
        new_table = deepcopy(self)

    
        # parse header
        col_indexes = {}
        if dynamic_col_headers:
            for row in rng.iter_rows(start=header_row):
                for cidx,col in enumerate(row):
                    if col.val:
                        col_indexes[col.val] = cidx
                break
        else:
            for cidx,field in enumerate(self._model.__fields__):
                col_indexes[field] = cidx
        

        # setup system model
        EntryModel = create_model(
            self._model.__class__.__name__,
            id=uuid4(),
            sheet_row=0,
            sheet_addr="",
            sheet="",
            workbook=Path(),
            __base__=self._model,
        )


        # parse data
        for row_idx, row in enumerate(rng.iter_rows(start=header_row+1)):
            try:
                entry_dict = {k:row[idx].val for k,idx in col_indexes.items()}

                # --
                if include_sys_params:
                    entry_dict['id']=uuid4()
                    entry_dict['sheet_row']=row_idx
                    entry_dict['sheet_addr']=rng._address
                    entry_dict['sheet']=rng._sheet
                    entry_dict['workbook']=rng._workbook

                    new_table.add(data=entry_dict, model=EntryModel)
                else:
                    new_table.add(data=entry_dict)

            except Exception as e:
                # import traceback
                # traceback.print_exc()
                if not skip_nulls:
                    raise e

        return new_table
    
    def to_dataframe(self) -> pd.DataFrame:
        df = pd.DataFrame(
            columns=[name for name,field in self._model.__fields__.items()],
            data=[
                model.dict() for model in self.entries
            ]
        )
        return df

    def add(self, data: dict, model: BaseModel = None):
        if model:
            self.entries.append(model.parse_obj(data))
        else:
            self.entries.append(self._model.parse_obj(data))
        