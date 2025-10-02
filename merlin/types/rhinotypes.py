"""
rhinotypes.py
Rhino type imports and aliases for IronPython/CLR applications
"""

import clr

# Add references to Rhino assemblies
clr.AddReference("RhinoCommon")
clr.AddReference("Eto")

# Import Rhino namespaces
import Rhino
from Rhino.Geometry import *
from Rhino.DocObjects import *
from Rhino.DocObjects.Tables import LayerTable, ObjectTable

# Import Eto namespaces (UI framework)
import Eto
from Eto.Forms import *
from Eto.Drawing import *

# Type aliases matching C# using statements
# Geometry types
RhinoArc = Rhino.Geometry.Arc
RhinoBox = Rhino.Geometry.Box
RhinoBrep = Rhino.Geometry.Brep
RhinoCircle = Rhino.Geometry.Circle
RhinoCurve = Rhino.Geometry.Curve
RhinoEllipse = Rhino.Geometry.Ellipse
RhinoInterval = Rhino.Geometry.Interval
RhinoLine = Rhino.Geometry.Line
RhinoLineCurve = Rhino.Geometry.LineCurve
RhinoNurbsCurve = Rhino.Geometry.NurbsCurve
RhinoPlane = Rhino.Geometry.Plane
RhinoPoint2d = Rhino.Geometry.Point2d
RhinoPoint3d = Rhino.Geometry.Point3d
RhinoPolyCurve = Rhino.Geometry.PolyCurve
RhinoPolyLine = Rhino.Geometry.Polyline
RhinoRectangle3d = Rhino.Geometry.Rectangle3d
RhinoSurface = Rhino.Geometry.Surface
RhinoVector2d = Rhino.Geometry.Vector2d
RhinoVector3d = Rhino.Geometry.Vector3d

# DocObjects types
RhinoLayerTable = LayerTable
RhinoObjectTable = ObjectTable

# Define what gets exported with 'from rhinotypes import *'
__all__ = [
    # Aliases
    'RhinoArc',
    'RhinoBox',
    'RhinoBrep',
    'RhinoCircle',
    'RhinoCurve',
    'RhinoEllipse',
    'RhinoInterval',
    'RhinoLayerTable',
    'RhinoLine',
    'RhinoLineCurve',
    'RhinoNurbsCurve',
    'RhinoObjectTable',
    'RhinoPlane',
    'RhinoPoint2d',
    'RhinoPoint3d',
    'RhinoPolyCurve',
    'RhinoPolyLine',
    'RhinoRectangle3d',
    'RhinoSurface',
    'RhinoVector2d',
    'RhinoVector3d',
]