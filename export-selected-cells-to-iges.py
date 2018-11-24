import FreeCAD
from SpreadSheetUtil import getSelectionRange
from ExportTo import exportSelectedCellsTo, IgesExportFormat

exportSelectedCellsTo(getSelectionRange(),
                      FreeCAD.activeDocument(),
                      IgesExportFormat())
