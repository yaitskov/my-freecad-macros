import os.path
import os
from shutil import rmtree
from SpreadSheetUtil import rowCol2Addr


class ExportFormat:
    def exportTo(obj, destPath):
        raise Exception("Not implemented")

    def fileType(self):
        raise Exception("Not implemented")


class StlExportFormat(ExportFormat):
    def exportTo(self, obj, destPath):
        obj.Shape.exportStl(destPath + '.' + self.fileType())

    def fileType(self):
        return 'stl'


class StpExportFormat(StlExportFormat):
    def fileType(self):
        return 'stp'


class ObjectsExporter:
    def __init__(self, doc, fileFormat):
        self._doc = doc
        self._fileFormat = fileFormat
        self._dir = self.dirName() + "-as-" + fileFormat.fileType()

    def dirName(self):
        return os.path.splitext(self._doc.FileName)[0]

    def recreateFolder(self, path):
        if os.path.exists(path):
            rmtree(path)
        os.makedirs(path)

    def exportObjectsTo(self, objects):
        self.recreateFolder(self._dir)
        for obj in objects:
            self.exportObject(obj)

    def exportObject(self, obj):
        self._fileFormat.exportTo(obj.ref, os.path.join(self._dir, obj.name))


class ObjBox:
    def __init__(self, obj):
        self.ref = obj
        self.name = obj.Label


class VisiblePartsVisitor:
    def findVisibleObjects(self, doc):
        result = []
        for obj in doc.Objects:
            if obj.ViewObject.Visibility:
                if not self._ifGroup(obj):
                    self.exportObjectIf(obj, result)
        return result

    def _ifGroup(self, obj):
        return obj.isDerivedFrom("App::DocumentObjectGroup")

    def exportObjectIf(self, obj, result):
        if not obj.ViewObject.Visibility:
            return
        if hasattr(obj, 'Shape'):
            print("add object = %s" % obj.Label)
            result.append(ObjBox(obj))


class MultiObjBox(ObjBox):
    def __init__(self, obj, n):
        ObjBox.__init__(self, obj)
        self.name = "%s-%d" % (self.name, n)


class SelectedInSpreadSheetObjVisitor:
    def __init__(self, doc):
        self._doc = doc

    def _getCell(self, row, col):
        return self._doc.Spreadsheet.get(rowCol2Addr(row, col))

    def find(self, colRange):
        result = []
        for row in xrange(colRange.start.row, colRange.end.row + 1):
            objName = self._getCell(row, colRange.start.col)
            params = (rowCol2Addr(row, colRange.start.col), objName)
            if not objName or isinstance(objName, float):
                print("Skip cell %s = [%s]" % params)
                continue

            for obj in self._doc.getObjectsByLabel(objName):
                if obj is None:
                    continue
                copies = self._getCell(row, colRange.start.col + 1)
                if not isinstance(copies, float):
                    print("Skip %s = [%s]: copies is not float" % params)
                    continue
                copies = int(copies)
                for i in xrange(0, copies):
                    result.append(MultiObjBox(obj, i + 1))
        return result


def exportSelectedCellsTo(selRange, doc, exFormat):
    if selRange is None:
        print("No selected rows")
        return

    visitor = SelectedInSpreadSheetObjVisitor(doc)
    exporter = ObjectsExporter(doc, exFormat)
    exporter.exportObjectsTo(visitor.find(selRange))
