"""
cadtypes.py
AutoCAD type imports and aliases for IronPython/CLR applications
"""

import clr

# Add references to AutoCAD assemblies
clr.AddReference("acdbmgd")
clr.AddReference("acmgd")
clr.AddReference("accoremgd")

# Import AutoCAD namespaces
from Autodesk.AutoCAD.DatabaseServices import *
from Autodesk.AutoCAD.Geometry import *
from Autodesk.AutoCAD.GraphicsInterface import *
from Autodesk.AutoCAD.ApplicationServices import Document
from Autodesk.AutoCAD.ApplicationServices.Core import Application
from Autodesk.AutoCAD.EditorInput import Editor

# Type aliases matching C# using statements
# ApplicationServices
AcadApplication = Application
AcadDocument = Document
AcadEditor = Editor

# DatabaseServices types
AcadArc = Arc
AcadCircle = Circle
AcadDatabase = Database
AcadEllipse = Ellipse
AcadFace = Face
AcadFaceRecord = FaceRecord
AcadLayout = Layout
AcadLine = Line
AcadMaterial = Material
AcadObjectId = ObjectId
AcadPlotSettings = PlotSettings
AcadPolyFaceMesh = PolyFaceMesh
AcadPolyMeshType = PolyMeshType
AcadPolyFaceMeshVertex = PolyFaceMeshVertex
AcadPolygonMesh = PolygonMesh
AcadPolygonMeshVertex = PolygonMeshVertex
AcadPolyline = Polyline
AcadPolyline2d = Polyline2d
AcadPolyline3d = Polyline3d
AcadSolid = Solid
AcadSolid3d = Solid3d
AcadSpline = Spline
AcadSurface = Surface
AcadViewport = Viewport

# Geometry types
AcadCircularArc2d = CircularArc2d
AcadCubicSplineCurve2d = CubicSplineCurve2d
AcadCurve2d = Curve2d
AcadEllipticalArc2d = EllipticalArc2d
AcadLine2d = Line2d
AcadLineSegment2d = LineSegment2d
AcadNurbCurve2d = NurbCurve2d
AcadOffsetCurve2d = OffsetCurve2d
AcadPlane = Plane
AcadPoint2d = Point2d
AcadPoint3d = Point3d
AcadPoint2dCollection = Point2dCollection
AcadPoint3dCollection = Point3dCollection
AcadPolylineCurve2d = PolylineCurve2d
AcadSplineEntity2d = SplineEntity2d
AcadTolerance = Tolerance
AcadVector2d = Vector2d
AcadVector3d = Vector3d

# Define what gets exported with 'from cadtypes import *'
__all__ = [
    # Aliases
    'AcadApplication',
    'AcadArc',
    'AcadCircle',
    'AcadCircularArc2d',
    'AcadCubicSplineCurve2d',
    'AcadCurve2d',
    'AcadDatabase',
    'AcadDocument',
    'AcadEditor',
    'AcadEllipse',
    'AcadEllipticalArc2d',
    'AcadFace',
    'AcadFaceRecord',
    'AcadLayout',
    'AcadLine',
    'AcadLine2d',
    'AcadLineSegment2d',
    'AcadMaterial',
    'AcadNurbCurve2d',
    'AcadObjectId',
    'AcadOffsetCurve2d',
    'AcadPlane',
    'AcadPlotSettings',
    'AcadPoint2d',
    'AcadPoint3d',
    'AcadPoint2dCollection',
    'AcadPoint3dCollection',
    'AcadPolyFaceMesh',
    'AcadPolyMeshType',
    'AcadPolyFaceMeshVertex',
    'AcadPolygonMesh',
    'AcadPolygonMeshVertex',
    'AcadPolyline',
    'AcadPolyline2d',
    'AcadPolyline3d',
    'AcadPolylineCurve2d',
    'AcadSolid',
    'AcadSolid3d',
    'AcadSpline',
    'AcadSplineEntity2d',
    'AcadSurface',
    'AcadTolerance',
    'AcadVector2d',
    'AcadVector3d',
    'AcadViewport',
]
