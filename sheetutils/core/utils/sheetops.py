
from openpyxl.utils import get_column_letter



# =================== HELPERS =================== #
def _rc_to_addr(row, column, zero_indexed=True):
    """ moddified from Michael Currie's answer:
        https://stackoverflow.com/questions/31420817/convert-excel-row-column-indices-to-alphanumeric-cell-reference-in-python-openpy
    """
    if not isinstance(row,int):
        raise Exception("row must be an integer")

    if not isinstance(column,int):
        raise Exception("column must be an integer")

    if zero_indexed:
        row += 1
        column += 1

    return get_column_letter(column) + str(row)

def _col_to_num(col):
    """ moddified from Sylvain's answer:
        https://stackoverflow.com/questions/7261936/convert-an-excel-or-spreadsheet-column-letter-to-its-number-in-pythonic-fashion
    """
    num = 0
    if isinstance(col,str) and col:
        for c in col:
            if c.isalpha():
                num = num * 26 + (ord(c.upper()) - ord('A')) + 1
            else:
                raise Exception("Invalid column letter")
        return num
    else:
        raise Exception("invalid input type" )

def _addr_to_tuple(addr):
    """ returns the row and column as a tuple of integers from cell address 
    
    example:
        address="A14" --> (14, 0)
    """
    col_addr = ""
    row_addr = ""
    scan_idx = 0

    # find and process column
    if len(addr) <= 1 or addr.isalpha():
        raise Exception("Address must have at least one letter and 1 digit")

    if not addr[scan_idx].isalpha():
        raise Exception("address column should be alpha character")

    while scan_idx < len(addr):
        char = addr[scan_idx]
        if char.isalpha():
            col_addr += char
        else:
            break
        scan_idx += 1
    

    # find and process row
    if not addr[scan_idx].isdigit():
        raise Exception("address column should be a digit!")

    while scan_idx < len(addr):
        char = addr[scan_idx]
        if char.isdigit():
            row_addr += char
        else:
            break
        scan_idx += 1 

    return ( _col_to_num(col_addr), int(row_addr) )