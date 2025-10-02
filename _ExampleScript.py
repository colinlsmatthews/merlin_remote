import clr
import sys
import os

import clr
from System.Reflection import Assembly

executing_assembly = Assembly.GetExecutingAssembly()
version = executing_assembly.GetName().Version

roaming_folder = os.getenv('APPDATA')

sys.path.append(r"D:\Programs\Autodesk\AutoCAD 2024")
sys.path.append(roaming_folder + r"\Autodesk\ApplicationPlugins\AWI Rhino.Inside AutoCAD.bundle" + str(version) + r"\Win64")

clr.AddReference("AWI.RhinoInside.Services")
from AWI.RhinoInside.Services import *

clr.AddReference("AWI.RhinoInside.Interop")
import AWI
from AWI.RhinoInside.Interop.Geometry import *
from AWI.RhinoInside.Interop import *

clr.AddReference("AWI.RhinoInside.Core")
from AWI.RhinoInside.Core import *
from AWI.RhinoInside.Core.Interfaces import *

# AutoCAD File API
clr.AddReference("Acdbmgd")
import Autodesk
from Autodesk.AutoCAD.DatabaseServices import Curve as CadDBCurve
from Autodesk.AutoCAD.DatabaseServices import DBObjectCollection

# AutoCAD Application API
clr.AddReference("accoremgd")
from Autodesk.AutoCAD.ApplicationServices.Core import Application

import System
from System import Func

import Rhino.Geometry as rg

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
internalGeometryConverter = InternalGeometryConverter.Instance;

# For converting Rhino geometry types to AutoCAD geometry types or visa versa
rhinoGeometryConverter = RhinoGeometryConverter.Instance;

def CreateCADRectangle(transactionManagerWrapper):
    xyPlane = rg.Plane.WorldXY

    # The command internalUnits = mm therefore these units are in mm.
    rectangle = rg.Rectangle3d(xyPlane, 1500.0, 4300.0)

    rectangleCurve = rectangle.ToPolyline().ToPolylineCurve()

    # Conversion methods ending Db convert to curves that can be added to the document
    # Non-db methods create curves that cannot be added to the document. Units Conversion
    # is handled dynamically. 
    autoCADCurves = rhinoGeometryConverter.ConvertDb(rectangleCurve)

    # Get the model space block table to add the curves to and unwrap it.
    modelSpace = InteropConverter.Unwrap(transactionManagerWrapper.GetModelSpaceBlockTableRecord(True));
    
    # Unwrap the transaction manager
    transactionManger = InteropConverter.Unwrap(transactionManagerWrapper)

    for dbCurve in autoCADCurves:
        # Append the curve to the model space.
        modelSpace.AppendEntity(dbCurve);

        # Add the curve to the document DB.
        transactionManger.AddNewlyCreatedDBObject(dbCurve, True);

        # Call dispose on any AutoCAD object that implements IDisposable to avoid memory leaks.
        dbCurve.Dispose()

    return True

# Create a .NET Func<ITransactionManager, T> delegate to pass into Document.Transaction.
net_CreateCADRectangle = Func[ITransactionManager, bool](CreateCADRectangle)

# Run the transaction to modify the document.
success = document.Transaction(net_CreateCADRectangle)