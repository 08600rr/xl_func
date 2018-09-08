#!/usr/bin/env python3
#
# Copyright 2018 Jacky Leung 08_dot_600_rr_at_g_mail_dot_com
#
# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 2 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, see <http://www.gnu.org/licenses/>.

import unittest

"""
The function, cellname2colrow converts cell name
to the column and row numbers use in xlrd.

Usage::
    >>> import xl_func
    >>> col, row = xl_func.cellname2colrow('G10')

:param name: Cell name which human will use in M$ Excel.
:rtype: col, row : column and row integers
"""

def colname2num(name):
    col = 0
    length = len(name)
    for idx in range(length):
        letter = name[length-idx-1]
        col += (ord(letter.upper()) - ord('A') + 1) * (26**idx)
    return col-1

def cellname2colrow(name):
    for i in range(1, len(name)):
        if name[i:].isdecimal():
            return colname2num(name[:i]), int(name[i:])-1

class Testcellname2num(unittest.TestCase):

    def test_colname2num(self):
        self.assertEqual(colname2num('A'), 0)
        self.assertEqual(colname2num('Z'), 25)
        self.assertEqual(colname2num('AA'), 26)
        self.assertEqual(colname2num('AZ'), 51)
        self.assertEqual(colname2num('BA'), 52)
        self.assertEqual(colname2num('ZZ'), 701)
        self.assertEqual(colname2num('AAA'), 702)
        self.assertEqual(colname2num('ZZZ'), 18277)
        self.assertEqual(colname2num('XFD'), 16383)

    def test_cellname2colrow(self):
        self.assertEqual(cellname2colrow('A1'), (0, 0))
        self.assertEqual(cellname2colrow('Z01'), (25, 0))
        self.assertEqual(cellname2colrow('Z10'), (25, 9))
        self.assertEqual(cellname2colrow('AA100'), (26, 99))
        self.assertEqual(cellname2colrow('ZZ1000'), (701, 999))
        self.assertEqual(cellname2colrow('XFD1'), (16383, 0))
        self.assertEqual(cellname2colrow('z1'), (25, 0))

if __name__ == '__main__':
    unittest.main()
