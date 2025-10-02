"""
awitypes.py
AWI RhinoInside type imports and aliases for IronPython/CLR applications
"""

import clr

# Add reference to AWI RhinoInside assemblies
clr.AddReference("AWI.RhinoInside.Core")
clr.AddReference("AWI.RhinoInside.Interop")

# Import AWI namespaces
from AWI.RhinoInside.Core import *
from AWI.RhinoInside.Core.Interfaces import *
from AWI.RhinoInside.Interop import *
from AWI.RhinoInside.Interop.Geometry import *

# Import specific types for aliasing
from AWI.RhinoInside.Core import UnitSystem
from AWI.RhinoInside.Interop.Geometry import (
    LineCurve,
    Plane,
    Point3d,
    PolyCurve,
    Surface,
    Vector3d
)

# Type aliases matching C# using statements
AwiLineCurve = LineCurve
AwiPlane = Plane
AwiPoint3d = Point3d
AwiPolyCurve = PolyCurve
AwiSurface = Surface
AwiUnitSystem = UnitSystem
AwiVector3d = Vector3d

# Define what gets exported with 'from awitypes import *'
__all__ = [
    # Aliases
    'AwiLineCurve',
    'AwiPlane',
    'AwiPoint3d',
    'AwiPolyCurve',
    'AwiSurface',
    'AwiUnitSystem',
    'AwiVector3d',
]