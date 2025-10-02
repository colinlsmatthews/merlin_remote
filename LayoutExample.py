import clr
import sys
import os
import math
sys.path.append(r"\\gilns010\Merlin\PythonScripts")
from merlin import VERSION

from System.Reflection import Assembly

executing_assembly = Assembly.GetExecutingAssembly()

roaming_folder = os.getenv("APPDATA")

bundle_path = os.path.join(
    roaming_folder,
    "Autodesk",
    "ApplicationPlugins",
    "AWI Rhino.Inside AutoCAD.bundle",
    # For some reason `str(version)` returns 1.3.1.0 when the version is actually 1.0.1.0
    # For now version is supplied via external module, but this needs to be looked into
    VERSION,
    "Win64",
)

sys.path.append(r"C:\Programs\Autodesk\AutoCAD 2024")
sys.path.append(bundle_path)

# AutoCAD File API
clr.AddReference("Acdbmgd")
import Autodesk
from Autodesk.AutoCAD.DatabaseServices import Curve as CadDBCurve
from Autodesk.AutoCAD.DatabaseServices import DBObjectCollection
from Autodesk.AutoCAD.DatabaseServices import Arc as CadArc
from Autodesk.AutoCAD.DatabaseServices import Circle as CadCircle
from Autodesk.AutoCAD.DatabaseServices import Line as CadLine
from Autodesk.AutoCAD.Geometry import Point3d as CadPoint
from Autodesk.AutoCAD.Geometry import Vector3d as CadVector3d

# AutoCAD Application API
clr.AddReference("accoremgd")
from Autodesk.AutoCAD.ApplicationServices.Core import Application


clr.AddReference("AWI.RhinoInside.Interop")
import AWI
from AWI.RhinoInside.Interop.Geometry import *
from AWI.RhinoInside.Interop import *

clr.AddReference("AWI.RhinoInside.Core")
from AWI.RhinoInside.Core import *
from AWI.RhinoInside.Core.Interfaces import *

clr.AddReference("AWI.RhinoInside.ObjectArxWrapper")
from AWI.RhinoInside.ObjectArxWrapper import *

clr.AddReference("AWI.RhinoInside.Services")
from AWI.RhinoInside.Services import *

import System
from System import Func

import Rhino.Geometry as rg
import rhinoscriptsyntax as rs

# Document interop
documentMananger = AWI.RhinoInside.Interop.DocumentManager.Instance
document = documentMananger.Document

# Use closeAction.SetCloseAction(CloseActionType) to:
# 1. Close the document without saving (default) input CloseActionType.Unmodified.
# 2. Close the document and save changes input CloseActionType.Save.
# 3. Close the document, save changes, and close the underlying AutoCAD document input CloseActionType.SaveAndClose.
# Alternatively, if you want to save the document after modifying it you can use document.Transaction(func, true)
# which sets the save flags to option 2.
closeAction = documentMananger.CloseAction

# For converting internal geometry types to Rhino geometry types or visa versa
internalGeometryConverter = InternalGeometryConverter.Instance
# For converting Rhino geometry types to AutoCAD geometry types or visa versa
rhinoGeometryConverter = RhinoGeometryConverter.Instance

def ConfigurePlotSettings(layout):
    return None