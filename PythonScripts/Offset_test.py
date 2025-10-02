import clr
import sys
import os

print("Script loaded successfully")

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
#import rhinoscriptsyntax as rs

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
    rectangle = rg.Rectangle3d(xyPlane, 1000.0, 5000.0)

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

#Here we go
clr.AddReference("AcMgd")
clr.AddReference("AcCoreMgd")
from Autodesk.AutoCAD.ApplicationServices import Application
from Autodesk.AutoCAD.EditorInput import PromptPointOptions, PromptDoubleOptions, PromptStatus, PromptKeywordOptions
from Autodesk.AutoCAD.EditorInput import PromptEntityOptions, PromptStatus

def prompt_user_choice():
    doc = Application.DocumentManager.MdiActiveDocument
    ed = doc.Editor

    opts = PromptKeywordOptions("\nChoose an option:")
    opts.Keywords.Add("Jungle")
    opts.Keywords.Add("Vitral")
    opts.Keywords.Add("Harvest")

    # Optional: set default
    opts.Keywords.Default = "Jungle"
    opts.AllowNone = False

    res = ed.GetKeywords(opts)

    if res.Status == PromptStatus.OK:
        selected = res.StringResult
        ed.WriteMessage("\nYou selected: {0}".format(selected))
        return selected
    else:
        ed.WriteMessage("\nSelection canceled.")
        return None

def get_existing_geometry():
    doc = Application.DocumentManager.MdiActiveDocument
    ed = doc.Editor

    opts = PromptEntityOptions("\nSelect a closed polyline or region:")
    res = ed.GetEntity(opts)

    if res.Status == PromptStatus.OK:
        ent_id = res.ObjectId
        ed.WriteMessage("\nSelected entity ID: {}".format(ent_id))
        return ent_id
    else:
        ed.WriteMessage("\nSelection canceled.")
        return None

def GetPointsFromACADPolyline(entityObjectId):
    """
    Extracts points from an AutoCAD polyline entity
    
    Args:
        entityObjectId: AutoCAD ObjectId of the polyline
        
    Returns:
        List of AutoCAD Point3d objects, or empty list if extraction fails

    Notes:
        This works and gives you the coordinate points, independent of object type.
    """
    acadPoints = []
    
    try:
        # Get the actual entity from the ObjectId
        acadEntity = entityObjectId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
        
        # Method 1: For Polyline3d - try Points collection
        if hasattr(acadEntity, 'Points') and acadEntity.Points is not None:
            for i in range(acadEntity.Points.Count):
                point = acadEntity.Points[i]
                acadPoints.append(point)
                print("Found point: " + str(point.X) + ", " + str(point.Y) + ", " + str(point.Z))
        
        # Method 2: For regular Polyline - try NumberOfVertices
        elif hasattr(acadEntity, 'NumberOfVertices'):
            numVertices = acadEntity.NumberOfVertices
            for i in range(numVertices):
                if hasattr(acadEntity, 'GetPoint3dAt'):
                    point = acadEntity.GetPoint3dAt(i)
                    acadPoints.append(point)
                    print("Found point: " + str(point.X) + ", " + str(point.Y) + ", " + str(point.Z))
                elif hasattr(acadEntity, 'GetPointAt'):
                    point = acadEntity.GetPointAt(i)
                    acadPoints.append(point)
                    print("Found point: " + str(point.X) + ", " + str(point.Y) + ", " + str(point.Z))
        
        # Method 3: Try iterating through vertex ObjectIds (for Polyline3d)
        if len(acadPoints) == 0:
            for vertexId in acadEntity:
                vertex = vertexId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                if hasattr(vertex, 'Position'):
                    point = vertex.Position
                    acadPoints.append(point)
                    print("Found point: " + str(point.X) + ", " + str(point.Y) + ", " + str(point.Z))
                vertex.Dispose()
        
        print("Total points extracted: " + str(len(acadPoints)))
    
    except Exception as e:
        print("Error extracting points: " + str(e))
    
    return acadPoints

def ConvertACADPolylineToRhinoPolyline(entityObjectId):
    """
    Complete function that extracts points from an AutoCAD polyline, 
    converts them to Rhino points, and creates a Rhino Polyline
    
    Args:
        entityObjectId: AutoCAD ObjectId of the polyline
        
    Returns:
        Rhino.Geometry.Polyline object, or None if conversion fails
    """
    try:
        # STEP 1: Get the actual entity from the ObjectId
        acadEntity = entityObjectId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
        
        # STEP 2: Extract AutoCAD points from the polyline entity
        acadPoints = []
        
        # Method 1: For Polyline3d - try Points collection
        if hasattr(acadEntity, 'Points') and acadEntity.Points is not None:
            for i in range(acadEntity.Points.Count):
                point = acadEntity.Points[i]
                acadPoints.append(point)
                print("Found AutoCAD point: " + str(point.X) + ", " + str(point.Y) + ", " + str(point.Z))
        
        # Method 2: For regular Polyline - try NumberOfVertices
        elif hasattr(acadEntity, 'NumberOfVertices'):
            numVertices = acadEntity.NumberOfVertices
            for i in range(numVertices):
                if hasattr(acadEntity, 'GetPoint3dAt'):
                    point = acadEntity.GetPoint3dAt(i)
                    acadPoints.append(point)
                    print("Found AutoCAD point: " + str(point.X) + ", " + str(point.Y) + ", " + str(point.Z))
                elif hasattr(acadEntity, 'GetPointAt'):
                    point = acadEntity.GetPointAt(i)
                    acadPoints.append(point)
                    print("Found AutoCAD point: " + str(point.X) + ", " + str(point.Y) + ", " + str(point.Z))
        
        # Method 3: Try iterating through vertex ObjectIds (for Polyline3d)
        if len(acadPoints) == 0:
            for vertexId in acadEntity:
                vertex = vertexId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                if hasattr(vertex, 'Position'):
                    point = vertex.Position
                    acadPoints.append(point)
                    print("Found AutoCAD point: " + str(point.X) + ", " + str(point.Y) + ", " + str(point.Z))
                vertex.Dispose()
        
        if len(acadPoints) < 2:
            print("Need at least 2 points to create a polyline. Found: " + str(len(acadPoints)))
            return None
        
        print("Total AutoCAD points extracted: " + str(len(acadPoints)))
        
        # STEP 3: Convert AutoCAD points to Rhino points
        rhinoPoints = []
        
        for i, acadPoint in enumerate(acadPoints):
            # Use the geometry converter to convert AutoCAD Point3d to Rhino Point3d
            rhinoPoint = rhinoGeometryConverter.Convert(acadPoint)
            if rhinoPoint is not None:
                rhinoPoints.append(rhinoPoint)
                print("Converted Rhino point " + str(i) + ": " + str(rhinoPoint.X) + ", " + str(rhinoPoint.Y) + ", " + str(rhinoPoint.Z))
            else:
                print("Failed to convert point " + str(i))
        
        if len(rhinoPoints) < 2:
            print("Point conversion failed - not enough valid Rhino points: " + str(len(rhinoPoints)))
            return None
        
        print("Total Rhino points converted: " + str(len(rhinoPoints)))
        
        # STEP 4: Create Rhino Polyline from the converted points
        rhinoPolyline = rg.Polyline(rhinoPoints)
        
        print("Successfully created Rhino Polyline with " + str(len(rhinoPoints)) + " points")
        
        return rhinoPolyline
        
    except Exception as e:
        print("Error in polyline conversion: " + str(e))
        return None

def OffsetRhinoPolylineCurveToCentroid(rhinoPolylineCurve, offsetDistance):
    """
    Takes a Rhino PolylineCurve and offsets it towards its centroid
    
    Args:
        rhinoPolylineCurve: Rhino.Geometry.PolylineCurve object
        offsetDistance: Distance to offset towards centroid (positive value)
        
    Returns:
        Rhino curve representing the offset polyline, or None if offset fails
    """
    try:
        # STEP 1: Ensure we have a valid curve
        if rhinoPolylineCurve is None:
            print("Invalid polyline curve provided")
            return None
        
        # STEP 2: Calculate the centroid (area center)
        if rhinoPolylineCurve.IsClosed:
            areaProperties = rg.AreaMassProperties.Compute(rhinoPolylineCurve)
            if areaProperties is not None:
                centroid = areaProperties.Centroid
                print("Calculated centroid: " + str(centroid.X) + ", " + str(centroid.Y) + ", " + str(centroid.Z))
            else:
                print("Could not calculate area centroid, using bounding box center")
                boundingBox = rhinoPolylineCurve.GetBoundingBox(rg.Plane.WorldXY)
                centroid = boundingBox.Center
        else:
            print("Curve is not closed, using bounding box center")
            boundingBox = rhinoPolylineCurve.GetBoundingBox(rg.Plane.WorldXY)
            centroid = boundingBox.Center
        
        # STEP 3: Offset the curve towards the centroid
        offsetPlane = rg.Plane.WorldXY
        inwardOffsetDistance = -abs(offsetDistance)
        
        offsetCurves = rhinoPolylineCurve.Offset(
            offsetPlane,
            inwardOffsetDistance,
            0.01,  # Tolerance
            rg.CurveOffsetCornerStyle.Sharp
        )
        
        if offsetCurves is None or len(offsetCurves) == 0:
            print("Offset operation failed")
            return None
        
        offsetCurve = offsetCurves[0]
        print("Successfully created offset curve")
        
        return offsetCurve
        
    except Exception as e:
        print("Error in polyline curve offset: " + str(e))
        return None

def build_pattern_in_region(transactionManagerWrapper):
    doc = Application.DocumentManager.MdiActiveDocument
    ed = doc.Editor

    # Step 1: Prompt user input
    pattern = prompt_user_choice()
    if pattern is None:
        return False
    region = get_existing_geometry()
    if region is None:
        return False
    
    rg_region = ConvertACADPolylineToRhinoPolyline(region)
    rg_region_offset = OffsetRhinoPolylineCurveToCentroid(rg_region, 1)

    print(pattern)
    print(region)
    print(rg_region)
    print(rg_region_offset)

    # Conversion methods ending Db convert to curves that can be added to the document
    # Non-db methods create curves that cannot be added to the document. Units Conversion
    # is handled dynamically. 
    autoCADCurves = rhinoGeometryConverter.ConvertDb(rg_region)

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
net_pattern_curves = Func[ITransactionManager, bool](build_pattern_in_region)

# Run the transaction to modify the document.
success = document.Transaction(net_pattern_curves)