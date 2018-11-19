import FreeCAD
from ExportTo import VisiblePartsVisitor, ObjectsExporter, StlExportFormat

doc = FreeCAD.activeDocument()
ObjectsExporter(doc, StlExportFormat()).exportObjectsTo(
    VisiblePartsVisitor().findVisibleObjects(doc))
