import FreeCAD
from SpreadSheetUtil import getSelectionRange
from ExportTo import exportSelectedCellsTo, StpExportFormat

exportSelectedCellsTo(getSelectionRange(),
                      FreeCAD.activeDocument(),
                      StpExportFormat())
