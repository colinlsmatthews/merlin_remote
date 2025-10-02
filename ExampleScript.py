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
internalGeometryConverter = InternalGeometryConverter.Instance;

# For converting Rhino geometry types to AutoCAD geometry types or visa versa
rhinoGeometryConverter = RhinoGeometryConverter.Instance;

from AWI.RhinoInside.Core import UnitSystem
# tell the converter your "internal" units are inches
internal_unit_system = UnitSystem.Inches

def CreateCADRectangle(transactionManagerWrapper):
    xyPlane = rg.Plane.WorldXY

    # The command internalUnits = mm therefore these units are in mm.
    rectangle = rg.Rectangle3d(xyPlane, 23.0, 84.0)

    rectangleCurve = rectangle.ToPolyline().ToPolylineCurve()

    # Conversion methods ending Db convert to curves that can be added to the document
    # Non-db methods create curves that cannot be added to the document. Units Conversion
    # is handled dynamically. 
    autoCADCurve = rhinoGeometryConverter.ConvertDb(rectangleCurve)


    # Call dispose on any AutoCAD object that implements IDisposable to avoid memory leaks.
    autoCADCurve.Dispose()
    rectangleCurve.Dispose()

    return True

def CreateCADRectangleOLD(transactionManagerWrapper):
    xyPlane = rg.Plane.WorldXY

    # The command internalUnits = mm therefore these units are in mm.
    rectangle = rg.Rectangle3d(xyPlane, 1500.0, 4300.0)

    rectangleCurve = rectangle.ToPolyline().ToPolylineCurve()

    # Conversion methods ending Db convert to curves that can be added to the document
    # Non-db methods create curves that cannot be added to the document. Units Conversion
    # is handled dynamically. 
    autoCADCurve = rhinoGeometryConverter.ConvertDb(rectangleCurve)

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


def CreateCADSrf4pt(transactionManagerWrapper):
    # Create surface points
    points = [(0,0,0), (20,3,0), (25, 35,0), (-2,30,0)]
    
    # Convert to Point3d objects
    rhinoPoints = [rg.Point3d(x, y, z) for x, y, z in points]
    
    # Create surface from points using Rhino common geometry
    srfGeometry = rg.NurbsSurface.CreateFromCorners(rhinoPoints[0], rhinoPoints[1], rhinoPoints[2], rhinoPoints[3])
    
    # Convert surface to Brep for better CAD compatibility
    brepGeometry = srfGeometry.ToBrep()
    
    # Convert the Rhino brep geometry to AutoCAD
    autoCADSrf = rhinoGeometryConverter.ConvertDb(brepGeometry)

    # Get the model space block table to add the surfaces to and unwrap it.
    modelSpace = InteropConverter.Unwrap(transactionManagerWrapper.GetModelSpaceBlockTableRecord(True));
    
    # Unwrap the transaction manager
    transactionManger = InteropConverter.Unwrap(transactionManagerWrapper)

    for dbSrf in autoCADSrf:
        # Append the surface to the model space.
        modelSpace.AppendEntity(dbSrf);

        # Add the surface to the document DB.
        transactionManger.AddNewlyCreatedDBObject(dbSrf, True);

        # Call dispose on any AutoCAD object that implements IDisposable to avoid memory leaks.
        dbSrf.Dispose()  # Fixed: was dbCurve.Dispose(), now dbSrf.Dispose()

    return True

# Create a .NET Func<ITransactionManager, T> delegate to pass into Document.Transaction.
net_CreateCADSrf4pt = Func[ITransactionManager, bool](CreateCADSrf4pt)

# Run the transaction to modify the document.
#success = document.Transaction(net_CreateCADSrf4pt)

def CreateCADPolySurfaceBox(transactionManagerWrapper):
    """Create a simple box polysurface"""
    # Define box dimensions
    width = 30.0
    height = 20.0
    depth = 15.0
    
    # Create a box as a polysurface
    box = rg.Box(rg.Plane.WorldXY, rg.Interval(0, width), rg.Interval(0, height), rg.Interval(0, depth))
    brepGeometry = box.ToBrep()
    
    # Convert the Rhino brep geometry to AutoCAD
    autoCADPolySrf = rhinoGeometryConverter.ConvertDb(brepGeometry)

    # Get the model space block table to add the surfaces to and unwrap it.
    modelSpace = InteropConverter.Unwrap(transactionManagerWrapper.GetModelSpaceBlockTableRecord(True))
    
    # Unwrap the transaction manager
    transactionManger = InteropConverter.Unwrap(transactionManagerWrapper)

    for dbSrf in autoCADPolySrf:
        # Append the surface to the model space.
        modelSpace.AppendEntity(dbSrf)

        # Add the surface to the document DB.
        transactionManger.AddNewlyCreatedDBObject(dbSrf, True)

        # Call dispose on any AutoCAD object that implements IDisposable to avoid memory leaks.
        dbSrf.Dispose()

    return True

def CreateCADPolySurfaceCylinder(transactionManagerWrapper):
    """Create a cylindrical polysurface"""
    # Define cylinder parameters
    baseCenter = rg.Point3d(0, 0, 0)
    topCenter = rg.Point3d(0, 0, 25)
    radius = 10.0
    
    # Create cylinder
    cylinder = rg.Cylinder(rg.Circle(baseCenter, radius), 25.0)
    brepGeometry = cylinder.ToBrep(True, True)  # True for cap bottom and top
    
    # Convert the Rhino brep geometry to AutoCAD
    autoCADPolySrf = rhinoGeometryConverter.ConvertDb(brepGeometry)

    # Get the model space block table to add the surfaces to and unwrap it.
    modelSpace = InteropConverter.Unwrap(transactionManagerWrapper.GetModelSpaceBlockTableRecord(True))
    
    # Unwrap the transaction manager
    transactionManger = InteropConverter.Unwrap(transactionManagerWrapper)

    for dbSrf in autoCADPolySrf:
        # Append the surface to the model space.
        modelSpace.AppendEntity(dbSrf)

        # Add the surface to the document DB.
        transactionManger.AddNewlyCreatedDBObject(dbSrf, True)

        # Call dispose on any AutoCAD object that implements IDisposable to avoid memory leaks.
        dbSrf.Dispose()

    return True

def CreateCADPolySurfaceFromSurfaces(transactionManagerWrapper, surfaces):
    """Create a polysurface from multiple individual surfaces"""
    # Convert surface list to breps
    breps = []
    for srf in surfaces:
        if hasattr(srf, 'ToBrep'):
            breps.append(srf.ToBrep())
        else:
            breps.append(srf)  # Already a brep
    
    # Join the breps into a polysurface
    joinedBreps = rg.Brep.JoinBreps(breps, 0.01)  # 0.01 is tolerance
    
    if joinedBreps and len(joinedBreps) > 0:
        polysurfaceBrep = joinedBreps[0]  # Take the first joined result
    else:
        # If joining fails, create a compound brep
        polysurfaceBrep = rg.Brep()
        for brep in breps:
            polysurfaceBrep.Append(brep)
    
    # Convert the Rhino brep geometry to AutoCAD
    autoCADPolySrf = rhinoGeometryConverter.ConvertDb(polysurfaceBrep)

    # Get the model space block table to add the surfaces to and unwrap it.
    modelSpace = InteropConverter.Unwrap(transactionManagerWrapper.GetModelSpaceBlockTableRecord(True))
    
    # Unwrap the transaction manager
    transactionManger = InteropConverter.Unwrap(transactionManagerWrapper)

    for dbSrf in autoCADPolySrf:
        # Append the surface to the model space.
        modelSpace.AppendEntity(dbSrf)

        # Add the surface to the document DB.
        transactionManger.AddNewlyCreatedDBObject(dbSrf, True)

        # Call dispose on any AutoCAD object that implements IDisposable to avoid memory leaks.
        dbSrf.Dispose()

    return True

def CreateCADPolySurfaceCustom(transactionManagerWrapper):
    """Create a custom polysurface from multiple surfaces"""
    # Create multiple surfaces that will form a polysurface
    surfaces = []
    
    # Bottom surface
    bottomPoints = [rg.Point3d(0,0,0), rg.Point3d(20,0,0), rg.Point3d(20,20,0), rg.Point3d(0,20,0)]
    bottomSrf = rg.NurbsSurface.CreateFromCorners(bottomPoints[0], bottomPoints[1], bottomPoints[2], bottomPoints[3])
    surfaces.append(bottomSrf)
    
    # Top surface (offset up)
    topPoints = [rg.Point3d(0,0,15), rg.Point3d(20,0,15), rg.Point3d(20,20,15), rg.Point3d(0,20,15)]
    topSrf = rg.NurbsSurface.CreateFromCorners(topPoints[0], topPoints[1], topPoints[2], topPoints[3])
    surfaces.append(topSrf)
    
    # Side surfaces
    # Front
    frontSrf = rg.NurbsSurface.CreateFromCorners(bottomPoints[0], bottomPoints[1], topPoints[1], topPoints[0])
    surfaces.append(frontSrf)
    
    # Right
    rightSrf = rg.NurbsSurface.CreateFromCorners(bottomPoints[1], bottomPoints[2], topPoints[2], topPoints[1])
    surfaces.append(rightSrf)
    
    # Back
    backSrf = rg.NurbsSurface.CreateFromCorners(bottomPoints[2], bottomPoints[3], topPoints[3], topPoints[2])
    surfaces.append(backSrf)
    
    # Left
    leftSrf = rg.NurbsSurface.CreateFromCorners(bottomPoints[3], bottomPoints[0], topPoints[0], topPoints[3])
    surfaces.append(leftSrf)
    
    # Convert surfaces to breps and join them
    breps = [srf.ToBrep() for srf in surfaces]
    joinedBreps = rg.Brep.JoinBreps(breps, 0.01)
    
    if joinedBreps and len(joinedBreps) > 0:
        polysurfaceBrep = joinedBreps[0]
    else:
        # Create a single brep from all surfaces
        polysurfaceBrep = rg.Brep()
        for brep in breps:
            for face in brep.Faces:
                polysurfaceBrep.Faces.Add(face)
    
    # Convert the Rhino brep geometry to AutoCAD
    autoCADPolySrf = rhinoGeometryConverter.ConvertDb(polysurfaceBrep)

    # Get the model space block table to add the surfaces to and unwrap it.
    modelSpace = InteropConverter.Unwrap(transactionManagerWrapper.GetModelSpaceBlockTableRecord(True))
    
    # Unwrap the transaction manager
    transactionManger = InteropConverter.Unwrap(transactionManagerWrapper)

    for dbSrf in autoCADPolySrf:
        # Append the surface to the model space.
        modelSpace.AppendEntity(dbSrf)

        # Add the surface to the document DB.
        transactionManger.AddNewlyCreatedDBObject(dbSrf, True)

        # Call dispose on any AutoCAD object that implements IDisposable to avoid memory leaks.
        dbSrf.Dispose()

    return True

# Create .NET Func delegates for each function
net_CreateCADPolySurfaceBox = Func[ITransactionManager, bool](CreateCADPolySurfaceBox)
net_CreateCADPolySurfaceCylinder = Func[ITransactionManager, bool](CreateCADPolySurfaceCylinder)
net_CreateCADPolySurfaceCustom = Func[ITransactionManager, bool](CreateCADPolySurfaceCustom)

#To Activate:
#success = document.Transaction(net_CreateCADPolySurfaceBox)
#success = document.Transaction(net_CreateCADPolySurfaceCylinder)
#success = document.Transaction(net_CreateCADPolySurfaceCustom)

# For using the function that takes parameters, you'd need to create a wrapper:
def CreatePolySurfaceFromSurfacesWrapper(surfaces_list):
    def wrapper(transactionManagerWrapper):
        return CreateCADPolySurfaceFromSurfaces(transactionManagerWrapper, surfaces_list)
    return wrapper

# Example usage with custom surfaces:
# my_surfaces = [surface1, surface2, surface3]  # Your list of surfaces
# wrapper_func = CreatePolySurfaceFromSurfacesWrapper(my_surfaces)
# net_wrapper = Func[ITransactionManager, bool](wrapper_func)
# success = document.Transaction(net_wrapper)

####################################################

def CreateCADLoftedSurfaceSimple(transactionManagerWrapper, add_profile_curves=False):
    """
    Builds 4 rectangular polylines at different Z levels (in INTERNAL units),
    lofts them into a Brep surface, converts to AutoCAD DB entities, and appends
    them to ModelSpace under the active transaction.

    If add_profile_curves=True, it also drops the source profile curves into ModelSpace.
    """
    # Unwrap DB objects (ModelSpace opened for write)
    model_space = InteropConverter.Unwrap(transactionManagerWrapper.GetModelSpaceBlockTableRecord(True))
    transaction = InteropConverter.Unwrap(transactionManagerWrapper)

    # --- 1) Build UN-SCALED Rhino curves (units = your app's InternalUnitSystem)
    def _closed_polyline_curve(pts):
        pl = rg.Polyline(pts)
        pl.Add(pts[0])
        return pl.ToPolylineCurve()  # IDisposable

    loft_curve1 = _closed_polyline_curve([rg.Point3d( 0,  0,  0), rg.Point3d(20,  0,  0),
                                          rg.Point3d(20, 15,  0), rg.Point3d( 0, 15,  0)])
    loft_curve2 = _closed_polyline_curve([rg.Point3d( 2,  1.5, 10), rg.Point3d(18,  1.5, 10),
                                          rg.Point3d(18, 13.5, 10), rg.Point3d( 2, 13.5, 10)])
    loft_curve3 = _closed_polyline_curve([rg.Point3d( 4,  3, 20), rg.Point3d(16,  3, 20),
                                          rg.Point3d(16, 12, 20), rg.Point3d( 4, 12, 20)])
    loft_curve4 = _closed_polyline_curve([rg.Point3d( 6,  4.5, 30), rg.Point3d(14,  4.5, 30),
                                          rg.Point3d(14, 10.5, 30), rg.Point3d( 6, 10.5, 30)])

    loft_curves = [loft_curve1, loft_curve2, loft_curve3, loft_curve4]

    # --- 2) Loft to Brep
    lofted = rg.Brep.CreateFromLoft(loft_curves, rg.Point3d.Unset, rg.Point3d.Unset,
                                    rg.LoftType.Normal, closed=False)
    if not lofted or len(lofted) == 0:
        # Clean up curves if loft failed
        for c in loft_curves:
            try: c.Dispose()
            except: pass
        print("[CreateCADLoftedSurfaceSimple] Loft failed.")
        return False

    loft_brep = lofted[0]  # IDisposable

    cad_entities = None
    try:
        # --- 3) Convert lofted Brep to AutoCAD DB objects
        converter = RhinoGeometryConverter.Instance
        cad_entities = converter.ConvertDb(loft_brep)

        # Some converters return a single entity, others an iterable — normalize
        if cad_entities is None:
            print("[CreateCADLoftedSurfaceSimple] ConvertDb(loft_brep) returned None.")
            return False
        if not hasattr(cad_entities, "__iter__"):
            cad_entities = [cad_entities]

        # --- 4) Append + register
        for db_obj in cad_entities:
            obj_id = model_space.AppendEntity(db_obj)
            transaction.AddNewlyCreatedDBObject(db_obj, True)
            # Safe to dispose after registering with the transaction
            db_obj.Dispose()

        # (Optional) also drop the Rhino profile curves into ModelSpace (as curves)
        if add_profile_curves:
            for c in loft_curves:
                db_curve = converter.ConvertDb(c)
                if db_curve is None:
                    continue
                if not hasattr(db_curve, "__iter__"):
                    db_curve = [db_curve]
                for ent in db_curve:
                    obj_id = model_space.AppendEntity(ent)
                    transaction.AddNewlyCreatedDBObject(ent, True)
                    ent.Dispose()

        print("[CreateCADLoftedSurfaceSimple] Loft and (optional) profiles created.")
        return True

    finally:
        # --- 5) Dispose Rhino objects
        try:
            loft_brep.Dispose()
        except: pass
        for c in loft_curves:
            try: c.Dispose()
            except: pass

# Create .NET Func delegate
net_CreateCADLoftedSurfaceSimple = Func[ITransactionManager, bool](CreateCADLoftedSurfaceSimple)

# Usage:
#success = document.Transaction(net_CreateCADLoftedSurfaceSimple)

#########################################

def CreateCADUndulatingLoftedSurface(transactionManagerWrapper):
    # Get the model space block table and transaction manager
    modelSpace = InteropConverter.Unwrap(transactionManagerWrapper.GetModelSpaceBlockTableRecord(True))
    transactionManger = InteropConverter.Unwrap(transactionManagerWrapper)
    
    # Scale factor for curves only
    scale = 25.4  # Convert inches to mm for curves
    
    # STEP 1: Create ORIGINAL unscaled undulating curves using polylines for the lofted surface
    
    # Curve 1 - Bottom undulating curve
    points1_orig = [
        rg.Point3d(0, 0, 0),
        rg.Point3d(5, 3, 0),
        rg.Point3d(10, -2, 0),
        rg.Point3d(15, 4, 0),
        rg.Point3d(20, 1, 0),
        rg.Point3d(25, -1, 0),
        rg.Point3d(30, 2, 0)
    ]
    polyline1_orig = rg.Polyline(points1_orig)
    loftCurve1 = polyline1_orig.ToPolylineCurve()
    
    # Curve 2 - Second level - different undulation pattern
    points2_orig = [
        rg.Point3d(2, 1, 8),
        rg.Point3d(7, -2, 8),
        rg.Point3d(12, 4, 8),
        rg.Point3d(17, 0, 8),
        rg.Point3d(22, 3, 8),
        rg.Point3d(27, -1, 8),
        rg.Point3d(32, 2, 8)
    ]
    polyline2_orig = rg.Polyline(points2_orig)
    loftCurve2 = polyline2_orig.ToPolylineCurve()
    
    # Curve 3 - Third level - more pronounced undulation
    points3_orig = [
        rg.Point3d(4, -1, 16),
        rg.Point3d(9, 5, 16),
        rg.Point3d(14, -3, 16),
        rg.Point3d(19, 4, 16),
        rg.Point3d(24, 0, 16),
        rg.Point3d(29, 3, 16),
        rg.Point3d(34, 1, 16)
    ]
    polyline3_orig = rg.Polyline(points3_orig)
    loftCurve3 = polyline3_orig.ToPolylineCurve()
    
    # Curve 4 - Top curve - gentle waves
    points4_orig = [
        rg.Point3d(6, 2, 24),
        rg.Point3d(11, -1, 24),
        rg.Point3d(16, 3, 24),
        rg.Point3d(21, 1, 24),
        rg.Point3d(26, -2, 24),
        rg.Point3d(31, 2, 24),
        rg.Point3d(36, 0, 24)
    ]
    polyline4_orig = rg.Polyline(points4_orig)
    loftCurve4 = polyline4_orig.ToPolylineCurve()
    
    loftCurves = [loftCurve1, loftCurve2, loftCurve3, loftCurve4]
    
    # STEP 2: Create lofted surface from UNSCALED curves (this will be OPEN)
    loftedBreps = rg.Brep.CreateFromLoft(loftCurves, rg.Point3d.Unset, rg.Point3d.Unset, rg.LoftType.Normal, False)
    
    if loftedBreps and len(loftedBreps) > 0:
        loftedBrep = loftedBreps[0]
        
        # Convert the lofted surface to AutoCAD
        autoCADSrf = rhinoGeometryConverter.ConvertDb(loftedBrep)

        for dbSrf in autoCADSrf:
            # Append the surface to the model space
            modelSpace.AppendEntity(dbSrf)
            # Add the surface to the document DB
            transactionManger.AddNewlyCreatedDBObject(dbSrf, True)
            # Dispose of the surface
            dbSrf.Dispose()
    
    # STEP 3: Create and add SCALED curves as separate entities
    
    # Curve 1 - Bottom undulating curve (scaled)
    points1 = [rg.Point3d(x*scale, y*scale, z*scale) for x, y, z in [(0, 0, 0), (5, 3, 0), (10, -2, 0), (15, 4, 0), (20, 1, 0), (25, -1, 0), (30, 2, 0)]]
    polyline1 = rg.Polyline(points1)
    curve1 = polyline1.ToPolylineCurve()
    
    autoCADCurves1 = rhinoGeometryConverter.ConvertDb(curve1)
    for dbCurve in autoCADCurves1:
        modelSpace.AppendEntity(dbCurve)
        transactionManger.AddNewlyCreatedDBObject(dbCurve, True)
        dbCurve.Dispose()
    
    # Curve 2 - Second level (scaled)
    points2 = [rg.Point3d(x*scale, y*scale, z*scale) for x, y, z in [(2, 1, 8), (7, -2, 8), (12, 4, 8), (17, 0, 8), (22, 3, 8), (27, -1, 8), (32, 2, 8)]]
    polyline2 = rg.Polyline(points2)
    curve2 = polyline2.ToPolylineCurve()
    
    autoCADCurves2 = rhinoGeometryConverter.ConvertDb(curve2)
    for dbCurve in autoCADCurves2:
        modelSpace.AppendEntity(dbCurve)
        transactionManger.AddNewlyCreatedDBObject(dbCurve, True)
        dbCurve.Dispose()
    
    # Curve 3 - Third level (scaled)
    points3 = [rg.Point3d(x*scale, y*scale, z*scale) for x, y, z in [(4, -1, 16), (9, 5, 16), (14, -3, 16), (19, 4, 16), (24, 0, 16), (29, 3, 16), (34, 1, 16)]]
    polyline3 = rg.Polyline(points3)
    curve3 = polyline3.ToPolylineCurve()
    
    autoCADCurves3 = rhinoGeometryConverter.ConvertDb(curve3)
    for dbCurve in autoCADCurves3:
        modelSpace.AppendEntity(dbCurve)
        transactionManger.AddNewlyCreatedDBObject(dbCurve, True)
        dbCurve.Dispose()
    
    # Curve 4 - Top curve (scaled)
    points4 = [rg.Point3d(x*scale, y*scale, z*scale) for x, y, z in [(6, 2, 24), (11, -1, 24), (16, 3, 24), (21, 1, 24), (26, -2, 24), (31, 2, 24), (36, 0, 24)]]
    polyline4 = rg.Polyline(points4)
    curve4 = polyline4.ToPolylineCurve()
    
    autoCADCurves4 = rhinoGeometryConverter.ConvertDb(curve4)
    for dbCurve in autoCADCurves4:
        modelSpace.AppendEntity(dbCurve)
        transactionManger.AddNewlyCreatedDBObject(dbCurve, True)
        dbCurve.Dispose()

    return True

# Create .NET Func delegate for the working function
net_CreateCADUndulatingLoftedSurface = Func[ITransactionManager, bool](CreateCADUndulatingLoftedSurface)

# Usage:
#success = document.Transaction(net_CreateCADUndulatingLoftedSurface)

#############################################

def CreateCADHybridNurbsLoftedSurfaceWPolyLines(transactionManagerWrapper):
    # Get the model space block table and transaction manager
    modelSpace = InteropConverter.Unwrap(transactionManagerWrapper.GetModelSpaceBlockTableRecord(True))
    transactionManger = InteropConverter.Unwrap(transactionManagerWrapper)
    
    # Scale factor for curves only
    scale = 25.4  # Convert inches to mm for curves
    
    import math
    
    # STEP 1: Create ORIGINAL unscaled NURBS curves for the lofted surface
    
    # Curve 1 - Bottom undulating curve
    points1_orig = [
        rg.Point3d(0, 0, 0),
        rg.Point3d(5, 3, 0),
        rg.Point3d(10, -2, 0),
        rg.Point3d(15, 4, 0),
        rg.Point3d(20, 1, 0),
        rg.Point3d(25, -1, 0),
        rg.Point3d(30, 2, 0)
    ]
    loftCurve1 = rg.Curve.CreateInterpolatedCurve(points1_orig, 3)
    
    # Curve 2 - Second level
    points2_orig = [
        rg.Point3d(2, 1, 8),
        rg.Point3d(7, -2, 8),
        rg.Point3d(12, 4, 8),
        rg.Point3d(17, 0, 8),
        rg.Point3d(22, 3, 8),
        rg.Point3d(27, -1, 8),
        rg.Point3d(32, 2, 8)
    ]
    loftCurve2 = rg.Curve.CreateInterpolatedCurve(points2_orig, 3)
    
    # Curve 3 - Third level
    points3_orig = [
        rg.Point3d(4, -1, 16),
        rg.Point3d(9, 5, 16),
        rg.Point3d(14, -3, 16),
        rg.Point3d(19, 4, 16),
        rg.Point3d(24, 0, 16),
        rg.Point3d(29, 3, 16),
        rg.Point3d(34, 1, 16)
    ]
    loftCurve3 = rg.Curve.CreateInterpolatedCurve(points3_orig, 3)
    
    # Curve 4 - Top curve
    points4_orig = [
        rg.Point3d(6, 2, 24),
        rg.Point3d(11, -1, 24),
        rg.Point3d(16, 3, 24),
        rg.Point3d(21, 1, 24),
        rg.Point3d(26, -2, 24),
        rg.Point3d(31, 2, 24),
        rg.Point3d(36, 0, 24)
    ]
    loftCurve4 = rg.Curve.CreateInterpolatedCurve(points4_orig, 3)
    
    loftCurves = [loftCurve1, loftCurve2, loftCurve3, loftCurve4]
    
    # STEP 2: Create PURE NURBS lofted surface
    loftedBreps = rg.Brep.CreateFromLoft(loftCurves, rg.Point3d.Unset, rg.Point3d.Unset, rg.LoftType.Normal, False)
    
    if loftedBreps and len(loftedBreps) > 0:
        loftedBrep = loftedBreps[0]
        
        # Convert the NURBS surface directly (this works!)
        autoCADSrf = rhinoGeometryConverter.ConvertDb(loftedBrep)
        for dbSrf in autoCADSrf:
            modelSpace.AppendEntity(dbSrf)
            transactionManger.AddNewlyCreatedDBObject(dbSrf, True)
            dbSrf.Dispose()
    
    # STEP 3: Create SCALED curves as PolyCurves with smooth 3-degree segments
    
    all_points_orig = [points1_orig, points2_orig, points3_orig, points4_orig]
    #POLYLINE CONVERSION IS BELOW!!!!!!
    for points_orig in all_points_orig:
        # Scale the original control points
        import clr
        clr.AddReference("AcDbMgd")  # needed for Autodesk.AutoCAD.Geometry

        from System import Array
        from Autodesk.AutoCAD.Geometry import Point3d

        # Scale the original control points (assumes points_orig is a sequence of rg.Point3d)
        scaled_points = [rg.Point3d(p.X * scale, p.Y * scale, p.Z * scale) for p in points_orig]

        # Convert Rhino -> AutoCAD Point3d
        # If your converter returns AutoCAD Point3d already:
        scaled_cad_points_list = [
            rhinoGeometryConverter.Convert(pt) for pt in scaled_points
        ]
        # If your converter returns Rhino points instead, use this line instead of the one above:
        # scaled_cad_points_list = [Point3d(pt.X, pt.Y, pt.Z) for pt in scaled_points]

        # Optional: filter out any Nones from a converter
        scaled_cad_points_list = [p for p in scaled_cad_points_list if p is not None]

        # Create a typed .NET array Point3d[]
        scaled_cad_points = Array[Point3d](scaled_cad_points_list)

        # Create a smooth 3-degree curve through all points, then add to PolyCurve
        # polyCurve = rg.PolyCurve()
        
        # Create a single smooth 3-degree interpolated curve through all points
        # try:
        # Fallback to polyline if interpolated curve fails
        polytype = Autodesk.AutoCAD.DatabaseServices.Poly3dType.SimplePoly
        acadPoint3dCollection = Autodesk.AutoCAD.Geometry.Point3dCollection(scaled_cad_points)
        acadPolyline3d = Autodesk.AutoCAD.DatabaseServices.Polyline3d(polytype, acadPoint3dCollection, False)
            
        # except:
        #     smooth_curve = rg.Curve.CreateInterpolatedCurve(scaled_points, 3)
        #     if smooth_curve is not None:
        #         polyCurve.Append(smooth_curve)
       
        modelSpace.AppendEntity(acadPolyline3d)
        transactionManger.AddNewlyCreatedDBObject(acadPolyline3d, True)
        acadPolyline3d.Dispose()
    return True

# Create .NET Func delegate
CreateCADHybridNurbsLoftedSurfaceWPolyLines = Func[ITransactionManager, bool](CreateCADHybridNurbsLoftedSurfaceWPolyLines)

# Usage:
#success = document.Transaction(CreateCADHybridNurbsLoftedSurfaceWPolyLines)

#########################################

def CreateCADHybridNurbsLoftedSurface(transactionManagerWrapper, add_profile_curves=True):
    """
    Builds 4 smooth (degree-3) interpolated NURBS curves at different Z levels
    in the app's INTERNAL units, lofts them into a Brep, converts to AutoCAD DB
    objects, and appends them to ModelSpace under the active transaction.

    If add_profile_curves=True, also converts/appends the source NURBS curves.
    """
    # Unwrap DB objects (ModelSpace opened for write)
    model_space = InteropConverter.Unwrap(transactionManagerWrapper.GetModelSpaceBlockTableRecord(True))
    transaction = InteropConverter.Unwrap(transactionManagerWrapper)

    # --- 1) Build UN-SCALED NURBS curves (units = your InternalUnitSystem)
    points1 = [
        rg.Point3d(0, 0, 0),
        rg.Point3d(5, 3, 0),
        rg.Point3d(10, -2, 0),
        rg.Point3d(15, 4, 0),
        rg.Point3d(20, 1, 0),
        rg.Point3d(25, -1, 0),
        rg.Point3d(30, 2, 0)
    ]
    points2 = [
        rg.Point3d(2, 1, 8),
        rg.Point3d(7, -2, 8),
        rg.Point3d(12, 4, 8),
        rg.Point3d(17, 0, 8),
        rg.Point3d(22, 3, 8),
        rg.Point3d(27, -1, 8),
        rg.Point3d(32, 2, 8)
    ]
    points3 = [
        rg.Point3d(4, -1, 16),
        rg.Point3d(9, 5, 16),
        rg.Point3d(14, -3, 16),
        rg.Point3d(19, 4, 16),
        rg.Point3d(24, 0, 16),
        rg.Point3d(29, 3, 16),
        rg.Point3d(34, 1, 16)
    ]
    points4 = [
        rg.Point3d(6, 2, 24),
        rg.Point3d(11, -1, 24),
        rg.Point3d(16, 3, 24),
        rg.Point3d(21, 1, 24),
        rg.Point3d(26, -2, 24),
        rg.Point3d(31, 2, 24),
        rg.Point3d(36, 0, 24)
    ]

    # Degree-3 interpolated curves
    loft_curve1 = rg.Curve.CreateInterpolatedCurve(points1, 3)
    loft_curve2 = rg.Curve.CreateInterpolatedCurve(points2, 3)
    loft_curve3 = rg.Curve.CreateInterpolatedCurve(points3, 3)
    loft_curve4 = rg.Curve.CreateInterpolatedCurve(points4, 3)

    loft_curves = [loft_curve1, loft_curve2, loft_curve3, loft_curve4]

    # --- 2) Loft to Brep
    lofted = rg.Brep.CreateFromLoft(loft_curves, rg.Point3d.Unset, rg.Point3d.Unset,
                                    rg.LoftType.Normal, closed=False)
    if not lofted or len(lofted) == 0:
        for c in loft_curves:
            try: c.Dispose()
            except: pass
        print("[CreateCADHybridNurbsLoftedSurface] Loft failed.")
        return False

    loft_brep = lofted[0]  # IDisposable

    try:
        converter = RhinoGeometryConverter.Instance

        # --- 3) Convert lofted Brep to AutoCAD DB objects
        cad_entities = converter.ConvertDb(loft_brep)
        if cad_entities is None:
            print("[CreateCADHybridNurbsLoftedSurface] ConvertDb(loft_brep) returned None.")
            return False
        if not hasattr(cad_entities, "__iter__"):
            cad_entities = [cad_entities]

        # --- 4) Append + register loft result
        for db_obj in cad_entities:
            obj_id = model_space.AppendEntity(db_obj)
            transaction.AddNewlyCreatedDBObject(db_obj, True)
            db_obj.Dispose()

        # (Optional) also drop the Rhino profile curves into ModelSpace
        if add_profile_curves:
            for c in loft_curves:
                db_curve = converter.ConvertDb(c)
                if db_curve is None:
                    continue
                # Normalize return shape
                if not hasattr(db_curve, "__iter__"):
                    db_curve = [db_curve]
                for ent in db_curve:
                    obj_id = model_space.AppendEntity(ent)
                    transaction.AddNewlyCreatedDBObject(ent, True)
                    ent.Dispose()

        print("[CreateCADHybridNurbsLoftedSurface] NURBS loft and (optional) profiles created.")
        return True

    finally:
        # Dispose Rhino geometry
        try: loft_brep.Dispose()
        except: pass
        for c in loft_curves:
            try: c.Dispose()
            except: pass


# Create .NET Func delegate
net_CreateCADHybridNurbsLoftedSurface = Func[ITransactionManager, bool](CreateCADHybridNurbsLoftedSurface)

# Usage:
success = document.Transaction(net_CreateCADHybridNurbsLoftedSurface)