import FreeCAD
from SpreadSheetUtil import getSelectionRange
from ExportTo import exportSelectedCellsTo, StlExportFormat

exportSelectedCellsTo(getSelectionRange(),
                      FreeCAD.activeDocument(),
                      StlExportFormat())
