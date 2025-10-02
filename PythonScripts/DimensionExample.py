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
# sys.path.append(roaming_folder + r"\Autodesk\ApplicationPlugins\AWI Rhino.Inside AutoCAD.bundle" + str(version) + r"\Win64")
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

# NOTE
# For creating associated dimensions, we have to force AutoCAD to load
# the ObjectArx runtime extension that is created from the C++ ObjectArx
# project. Otherwise `clr.AddReference("AWI.RhinoInside.ObjectArxWrapper")`
# will throw an error and fail. The loading operation will show a warning
# in AutoCAD the first time and prompt the user to allow loading of a
# new extension. This might be confusing to some people.
arx_path = os.path.join(bundle_path, "AWI.RhinoInside.ObjectArx.arx")
Autodesk.AutoCAD.Runtime.SystemObjects.DynamicLinker.LoadModule(arx_path, True, False)

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


def CreateAssociatedAlignedExample(transactionManagerWrapper):
    """Example method to create an associated aligned dimension in the model space.

    Args:
        transactionManagerWrapper (ITransactionManager)
    """
    point1 = CadPoint(0, 0, 0)
    point2 = CadPoint(0, 50, 0)

    line1 = CadLine(point1, CadPoint(100, 0, 0))
    line2 = CadLine(point2, CadPoint(100, 50, 0))

    dimPoint = CadPoint(-10, 25, 0)
    
    # NOTE
    # the modelSpace and transactionManager instantiation works
    # differently in Python than in the C# examples. See below:
    modelSpace = InteropConverter.Unwrap(
        transactionManagerWrapper.GetModelSpaceBlockTableRecord(True)
    )
    transactionManager = InteropConverter.Unwrap(transactionManagerWrapper)

    lines = [line1, line2]

    createdIds = []

    for line in lines:
        id = modelSpace.AppendEntity(line)

        transactionManager.AddNewlyCreatedDBObject(line, True)

        createdIds.Add(ObjectId(id))

    pointObjectAssociation1 = PointObjectAssociation(
        internalGeometryConverter.Convert(point1), createdIds[0], 0
    )

    pointObjectAssociation2 = PointObjectAssociation(
        internalGeometryConverter.Convert(point2), createdIds[1], 0
    )

    dimensionPoint = internalGeometryConverter.Convert(dimPoint)

    associatedDimension = AssociatedAlignedDimension(
        pointObjectAssociation1, #IPointObjectAssociation startAssociation
        pointObjectAssociation2, #IPointObjectAssociation endAssociation
        dimensionPoint, #IPoint3d dimensionPoint
        ObjectId() #IObjectId dimensionStyle
        #[string textOverride = "<>"]
    )

    associatedDimension.AddToDocument()

    associatedDimension.Dispose()
    pointObjectAssociation1.Dispose()
    pointObjectAssociation2.Dispose()

# Create a .NET Func<ITransactionManager, T> delegate to pass into Document.Transaction.
net_CreateAssociatedAlignedExample = Func[ITransactionManager, bool](
    CreateAssociatedAlignedExample
)

# Run the transaction to modify the document.
success = document.Transaction(net_CreateAssociatedAlignedExample)


def CreateAssociatedRotatedExample(transactionManagerWrapper):
    """Example method to create an associated aligned dimension in the model space.

    Args:
        transactionManagerWrapper (ITransactionManager)
    """
    point1 = CadPoint(200, 0, 0)
    point2 = CadPoint(250, 50, 0)
    
    line1 = CadLine(point1, CadPoint(300, 0, 0))
    line2 = CadLine(point2, CadPoint(350, 50, 0))
    
    dimPoint = CadPoint(190, 25, 0)
    
    # NOTE
    # the modelSpace and transactionManager instantiation works
    # differently in Python than in the C# examples. See below:
    modelSpace = InteropConverter.Unwrap(
        transactionManagerWrapper.GetModelSpaceBlockTableRecord(True)
    )
    transactionManager = InteropConverter.Unwrap(transactionManagerWrapper)

    lines = [line1, line2]
    
    createdIds = []
    
    for line in lines:
        id = modelSpace.AppendEntity(line)
        
        transactionManager.AddNewlyCreatedDBObject(line, True)
        
        createdIds.Add(ObjectId(id))
        
        line.Dispose
    
    pointObjectAssociation1 = PointObjectAssociation(
        internalGeometryConverter.Convert(point1), createdIds[0], 0
    )
    
    pointObjectAssociation2 = PointObjectAssociation(
        internalGeometryConverter.Convert(point2), createdIds[1], 0
    )
    
    dimensionPoint = internalGeometryConverter.Convert(dimPoint)
    
    associatedDimension = AssociatedRotatedDimension(
        pointObjectAssociation1, #IPointObjectAssociation startAssociation
        pointObjectAssociation2, #IPointObjectAssociation endAssociation
        dimensionPoint, #IPoint3d dimensionPoint
        ObjectId(), #IObjectId dimensionStyle
        math.tan(50.0 / 50.0) #double rotation
        #[string textOverride = "<>"]
    )
    
    associatedDimension.AddToDocument()
    
    associatedDimension.Dispose()
    pointObjectAssociation1.Dispose()
    pointObjectAssociation2.Dispose()

# Create a .NET Func<ITransactionManager, T> delegate to pass into Document.Transaction.
net_CreateAssociatedRotatedExample = Func[ITransactionManager, bool](
    CreateAssociatedRotatedExample
)

# Run the transaction to modify the document.
success = document.Transaction(net_CreateAssociatedRotatedExample)


def CreateAssociatedArcExample(transactionManagerWrapper):
    center = CadPoint(500, 0, 0)
    radius = 50.0
    normal = CadVector3d(0, 0, 1)
    arc = CadArc(center, normal, radius, 0, math.pi / 2)
    dimPoint = CadPoint(575, 75, 0)

    # NOTE
    # the modelSpace and transactionManager instantiation works
    # differently in Python than in the C# examples. See below:
    modelSpace = InteropConverter.Unwrap(
        transactionManagerWrapper.GetModelSpaceBlockTableRecord(True)
    )
    transactionManager = InteropConverter.Unwrap(transactionManagerWrapper)
    
    id = modelSpace.AppendEntity(arc)
    
    objectId = ObjectId(id)
    
    transactionManager.AddNewlyCreatedDBObject(arc, True)
    
    centerPointObjectAssociation = PointObjectAssociation(
        internalGeometryConverter.Convert(center), objectId, 3
    )
    
    startPointObjectAssociation = PointObjectAssociation(
        internalGeometryConverter.Convert(arc.StartPoint), objectId, 1
    )
    
    endPointObjectAssociation = PointObjectAssociation(
        internalGeometryConverter.Convert(arc.EndPoint), objectId, 2
    )
    
    dimensionPoint = internalGeometryConverter.Convert(dimPoint)
    
    associatedDimension = AssociatedArcDimension(
        startPointObjectAssociation, #IPointObjectAssociation startAssociation
        endPointObjectAssociation, #IPointObjectAssociation endAssociation
        centerPointObjectAssociation, #IPointObjectAssociation centerAssociation
        dimensionPoint, #IPoint3d arcPoint
        ObjectId() #IObjectId dimensionStyle
        #[string textOverride = "<>"]
    )
    
    associatedDimension.AddToDocument()
    
    arc.Dispose()
    associatedDimension.Dispose()
    centerPointObjectAssociation.Dispose()
    startPointObjectAssociation.Dispose()
    endPointObjectAssociation.Dispose()
    
# Create a .NET Func<ITransactionManager, T> delegate to pass into Document.Transaction.
net_CreateAssociatedArcExample = Func[ITransactionManager, bool](
    CreateAssociatedArcExample
)

# Run the transaction to modify the document.
success = document.Transaction(net_CreateAssociatedArcExample)


def CreateAssociatedRadialExample(transactionManagerWrapper):
    radius = 50.0
    centerPoint = CadPoint(700, 0, 0)
    point2 = CadPoint(700 + radius, 0, 0)
    normal = CadVector3d(0, 0, 1)
    circle = CadCircle(centerPoint, normal, radius)
    
    # NOTE
    # the modelSpace and transactionManager instantiation works
    # differently in Python than in the C# examples. See below:
    modelSpace = InteropConverter.Unwrap(
        transactionManagerWrapper.GetModelSpaceBlockTableRecord(True)
    )
    transactionManager = InteropConverter.Unwrap(transactionManagerWrapper)

    id = modelSpace.AppendEntity(circle)
    objectId = ObjectId(id)
    
    transactionManager.AddNewlyCreatedDBObject(circle, True)
    circle.Dispose()
    
    pointObjectAssociation1 = PointObjectAssociation(
        internalGeometryConverter.Convert(centerPoint), objectId, 1
    )
    
    pointObjectAssociation2 = PointObjectAssociation(
        internalGeometryConverter.Convert(point2), objectId, 1
    )
    
    associatedDimension = AssociatedRadialDimension(
        pointObjectAssociation1, #IPointObjectAssociation centerAssociation
        pointObjectAssociation2, #IPointObjectAssociation cordAssociation
        ObjectId(), #IObjectId dimensionStyle
        20.0 #double leaderLength
        #[string textOverride = "<>"]
    )
    
    associatedDimension.AddToDocument()
    
    associatedDimension.Dispose()
    pointObjectAssociation1.Dispose()
    pointObjectAssociation2.Dispose()

# Create a .NET Func<ITransactionManager, T> delegate to pass into Document.Transaction.
net_CreateAssociatedRadialExample = Func[ITransactionManager, bool](
    CreateAssociatedRadialExample
)

# Run the transaction to modify the document.
success = document.Transaction(net_CreateAssociatedRadialExample)


def CreateAssociatedDiametricExample(transactionManagerWrapper):
    radius = 25.0
    centerPoint = CadPoint(1000, 0, 0)
    point1 = CadPoint(1000, -radius, 0)
    point2 = CadPoint(1000, radius, 0)
    normal = CadVector3d(0, 0, 1)
    circle = CadCircle(centerPoint, normal, radius)
    
    # NOTE
    # the modelSpace and transactionManager instantiation works
    # differently in Python than in the C# examples. See below:
    modelSpace = InteropConverter.Unwrap(
        transactionManagerWrapper.GetModelSpaceBlockTableRecord(True)
    )
    transactionManager = InteropConverter.Unwrap(transactionManagerWrapper)
    
    id = modelSpace.AppendEntity(circle)
    objectId = ObjectId(id)
    
    transactionManager.AddNewlyCreatedDBObject(circle, True)
    circle.Dispose()
    
    pointObjectAssociation1 = PointObjectAssociation(
        internalGeometryConverter.Convert(point1), objectId, 0
    )
    
    pointObjectAssociation2 = PointObjectAssociation(
        internalGeometryConverter.Convert(point2), objectId, 1
    )
    
    associatedDimension = AssociatedDiametricDimension(
        pointObjectAssociation1, #IPointObjectAssociation chordAssociation
        pointObjectAssociation2, #IPointObjectAssociation farChordAssociation
        ObjectId(), #IObjectId dimensionStyle
        10.0 #double leaderLength
        #[string textOverride = "<>"]
    )

    associatedDimension.AddToDocument()
    
    associatedDimension.Dispose()
    pointObjectAssociation1.Dispose()
    pointObjectAssociation2.Dispose()

# Create a .NET Func<ITransactionManager, T> delegate to pass into Document.Transaction.
net_CreateAssociatedDiametricExample = Func[ITransactionManager, bool](
    CreateAssociatedDiametricExample
)

# Run the transaction to modify the document.
success = document.Transaction(net_CreateAssociatedDiametricExample)


def CreateAssociatedAngularExample(transactionManagerWrapper):
    center = CadPoint(1200, 0, 0)
    point1 = CadPoint(1200, 50, 0)
    point2 = CadPoint(1250, 0, 0)
    line1 = CadLine(center, point1)
    line2 = CadLine(center, point2)
    dimPoint = CadPoint(1220, 25, 0)
    
    # NOTE
    # the modelSpace and transactionManager instantiation works
    # differently in Python than in the C# examples. See below:
    modelSpace = InteropConverter.Unwrap(
        transactionManagerWrapper.GetModelSpaceBlockTableRecord(True)
    )
    transactionManager = InteropConverter.Unwrap(transactionManagerWrapper)

    lines = [line1, line2]
    
    createdIds = []
    
    for line in lines:
        id = modelSpace.AppendEntity(line)
        transactionManager.AddNewlyCreatedDBObject(line, True)
        createdIds.Add(ObjectId(id))
        line.Dispose()
    
    centerPointObjectAssociation1 = PointObjectAssociation(
        internalGeometryConverter.Convert(center), createdIds[0], 3
    )
    
    pointObjectAssociation1 = PointObjectAssociation(
        internalGeometryConverter.Convert(point1), createdIds[0], 1
    )
    
    pointObjectAssociation2 = PointObjectAssociation(
        internalGeometryConverter.Convert(point2), createdIds[1], 2
    )
    
    dimensionPoint = internalGeometryConverter.Convert(dimPoint)
    
    associatedDimension = AssociatedAngularDimension(
        pointObjectAssociation1, #IPointObjectAssociation startAssociation
        pointObjectAssociation2, #IPointObjectAssociation endAssociation
        centerPointObjectAssociation1, #IPointObjectAssociation centerAssociation
        dimensionPoint, #IPoint3d arcPoint
        ObjectId() #IObjectId dimensionStyle
        #[string textOverride = "<>"]
    )

    associatedDimension.AddToDocument()
    
    associatedDimension.Dispose()
    centerPointObjectAssociation1.Dispose()
    pointObjectAssociation1.Dispose()
    pointObjectAssociation2.Dispose()

# Create a .NET Func<ITransactionManager, T> delegate to pass into Document.Transaction.
net_CreateAssociatedAngularExample = Func[ITransactionManager, bool](
    CreateAssociatedAngularExample
)

# Run the transaction to modify the document.
success = document.Transaction(net_CreateAssociatedAngularExample)