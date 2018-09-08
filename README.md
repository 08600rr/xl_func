# xl_func

The function, cellname2colrow converts cell name
to the column and row numbers use in xlrd.

Usage::
    >>> import xl_func
    >>> col, row = xl_func.cellname2colrow('G10')

:param name: Cell name which human will use in M$ Excel.
:rtype: col, row : column and row integers
