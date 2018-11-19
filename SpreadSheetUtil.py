# -*- coding: utf-8 -*-

import FreeCADGui
from PySide import QtGui

column_list = map(chr, range(ord('A'), ord('Z')))


class Point:
    def __init__(self, row, col):
        self.row = row
        self.col = col


def rowCol2Addr(row, col):
    return '%s%d' % (column_list[col], row)


class SelRange:
    def __init__(self, start, end):
        self.start = start
        self.end = end


def getSelectionRange():
        mw = FreeCADGui.getMainWindow()
        mdiarea = mw.findChild(QtGui.QMdiArea)

        subw = mdiarea.subWindowList()

        table = None
        for i in subw:
            if i.widget().metaObject().className() == "SpreadsheetGui::SheetView":
                sheet = i.widget()

                table = sheet.findChild(QtGui.QTableView)
        if table is None:
            print("No selection")
            return None
        ind = table.selectedIndexes()

        l_elements = len(ind)

        if l_elements == 0:
            print("No selection")
            return None

        first_element = ind.__getitem__(0)
        fe_col = first_element.column()
        fe_row = first_element.row() + 1

        last_element = ind.__getitem__(l_elements-1)
        le_col = last_element.column()
        le_row = last_element.row() + 1

        return SelRange(Point(fe_row, fe_col), Point(le_row, le_col))
