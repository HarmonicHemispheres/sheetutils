import sheetutils.core.utils.sheetops as shops
from typing import Any
from pydantic import BaseModel

class Cell(BaseModel):
    val: Any
    row: int = None
    col: int = None

    def addr(self):
        return shops._rc_to_addr(self.row, self.col)
        
