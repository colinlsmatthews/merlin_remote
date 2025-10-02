import sys
import math

sys.path.append(r"\\gilns010\Merlin\merlin_remote\PythonScripts")
from merlin import VERSION
from merlin.types import *

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
