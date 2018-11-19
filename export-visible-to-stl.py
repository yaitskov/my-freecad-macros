import FreeCAD
import os.path
import os
from shutil import rmtree


class ObjectsExporter:
    def __init__(self, doc):
        self._doc = doc
        self._dir = self.dirName() + "-" + self.fileType()

    def fileType(self):
        raise Exception("Not implemented")

    def dirName(self):
        return os.path.splitext(self._doc.FileName)[0]

    def recreateFolder(self, path):
        if os.path.exists(path):
            rmtree(path)
        os.makedirs(path)

    def exportVisibleObjectsTo(self):
        self.recreateFolder(self._dir)
        self.discoverVisibleObjects()

    def exportGroupObjects(self, group):
        for obj in group.OutList:
            if self._ifGroup(obj):
                self.exportGroupObject(obj)
            else:
                self.exportObjectIf(obj)

    def discoverVisibleObjects(self):
        for obj in self._doc.Objects:
            if obj.ViewObject.Visibility:
                if self._ifGroup(obj):
                    self.exportGroupObjects(obj)
                else:
                    self.exportObjectIf(obj)

    def _ifGroup(self, obj):
        return obj.isDerivedFrom("App::DocumentObjectGroup")

    def exportObjectIf(self, obj):
        if not obj.ViewObject.Visibility:
            return
        if hasattr(obj, 'Shape'):
            fname = obj.Label + "." + self.fileType()
            self.exportToFile(obj, os.path.join(self._dir, fname))

    def exportToFile(self, obj, destPath):
        raise Exception("Not implemented")


class ObjectsToStlExporter(ObjectsExporter):
    def fileType(self):
        return 'stl'

    def exportToFile(self, obj, destPath):
        obj.Shape.exportStl(destPath)


ObjectsToStlExporter(FreeCAD.activeDocument()).exportVisibleObjectsTo()
