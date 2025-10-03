import clr
import sys
import os
sys.path.append(r"\\gilns010\Merlin\merlin_remote\PythonScripts")
from merlin import VERSION
from merlin.types import *
from merlin.types.cadtypes import LoadModule

from System.Reflection import Assembly

roaming_folder = os.getenv("APPDATA")

bundle_path = os.path.join(
    roaming_folder,
    "Autodesk",
    "ApplicationPlugins",
    "AWI Rhino.Inside AutoCAD.bundle",
    VERSION,
    "Win64",
)

# AutoCAD File API
clr.AddReference("Acdbmgd")
import Autodesk

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
LoadModule(arx_path, True, False)

clr.AddReference("AWI.RhinoInside.ObjectArxWrapper")
from AWI.RhinoInside.ObjectArxWrapper import *

clr.AddReference("AWI.RhinoInside.Services")
from AWI.RhinoInside.Services import *

clr.AddReference("AWI.RhinoInside.API")
from AWI.RhinoInside.API import *

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

import uuid
import math
from typing import List, Optional


class SheetMetalFoldingResults:
    """Container class for sheet metal folding results"""
    def __init__(self):
        self.flat_curves: List = []
        self.fold_curves: List = []
        self.relief_curves: List = []
        self.labels: List[str] = []
        self.meshes: List = []
        self.breps: List = []


class SheetMetalFolder:
    """Main class for sheet metal folding operations"""
    
    def __init__(
        self,
        mold_model,
        flip: bool,
        material_thickness: float,
        material_k_factor: float,
        bend_radius: float,
        relief_definition,
        input_unit_system
    ):
        """
        Initialize the sheet metal folder
        
        Args:
            mold_model: RhinoBrep object representing the mold
            flip: Whether to flip the brep
            material_thickness: Thickness of the material
            material_k_factor: K-factor for bend calculations
            bend_radius: Radius of the bend
            relief_definition: Relief definition object
            input_unit_system: Unit system for input values
        """
        # Store constructor parameters as private instance attributes
        self._brep = mold_model
        self._flip = flip
        self._thickness = material_thickness
        self._k_factor = material_k_factor
        self._bend_radius = bend_radius
        self._relief_definition = relief_definition
        self._input_unit_system = input_unit_system
        
        # Internal unit system (constant)
        self._internal_unit_system = AwiUnitSystem.Inches
        
        # Unit system manager
        self._unit_system_manager = UnitSystemManager(
            self._input_unit_system, 
            self._internal_unit_system
        )
        
        # Geometry converters (initialized later)
        self._geometry_converter = None
        self._internal_geometry_converter = None
        
        # Axis definitions
        self._x_axis = AwiVector3d.x_axis()
        self._y_axis = AwiVector3d.y_axis()
        self._z_axis = AwiVector3d.z_axis()
        
        # Initialize converters
        self._initialize_converters(self._unit_system_manager)
    
    def fold(self) -> SheetMetalFoldingResults:
        """
        Main folding operation that calculates both 2D flat pattern and 3D folded model
        
        Returns:
            SheetMetalFoldingResults object containing all calculated geometry
        """
        # Initialize output container class
        output_results = SheetMetalFoldingResults()
        
        # Create material and tool definitions
        material_definition = SheetMetalMaterialDefinition()
        material_definition.id = uuid.uuid4()
        material_definition.thickness = UnitLength(self._thickness, self._input_unit_system)
        material_definition.name = "Grasshopper Dynamic Sheet Metal Material"
        material_definition.k_factor = self._k_factor
        
        folding_tool_definition = FoldingToolDefinition()
        folding_tool_definition.radius = UnitLength(self._bend_radius, self._input_unit_system)
        
        # Prepare brep for 2d and 3d calculations
        brep_for_2d = self._brep.duplicate_brep()
        if self._flip:
            brep_for_2d.flip()
        brep_for_3d = self._brep.duplicate_brep()
        
        # Calculate the 2d flat pattern
        solid = Solid(brep_for_2d)
        panel_mold_model_3d = PanelMoldModel3d(solid)
        folded_panel = FoldedPanel(
            panel_mold_model_3d, 
            material_definition, 
            folding_tool_definition
        )
        panel_flat = PanelFlat(folded_panel, self._relief_definition)
        
        drawing_plane_origin = Point3d(0, 0, 0)
        drawing_plane = Plane(drawing_plane_origin, self._x_axis, self._y_axis)
        panel_flat_drawing = PanelFlatDrawing(panel_flat, drawing_plane, True)
        
        # Calculate the 3d folded model
        solid_for_3d = Solid(brep_for_3d)
        panel_mold_model_3d_for_3d = PanelMoldModel3d(solid_for_3d)
        folded_panel_for_3d = FoldedPanel(
            panel_mold_model_3d_for_3d, 
            material_definition, 
            folding_tool_definition
        )
        model_3d = FoldedModel3d(folded_panel_for_3d)
        
        # Convert results to Rhino types
        self._create_panel_flat_in_rhino(panel_flat_drawing, output_results)
        self._create_folded_model_in_rhino(model_3d, output_results)
        
        return output_results
    
    def _create_panel_flat_in_rhino(
        self, 
        panel_flat_drawing, 
        results: SheetMetalFoldingResults
    ) -> None:
        """
        Create flat panel representation and populate results
        
        Args:
            panel_flat_drawing: IPanelFlatDrawing object
            results: SheetMetalFoldingResults object to populate
        """
        # Process panel outline curves
        flat_curves_output = []
        for outline_curve in panel_flat_drawing.panel_outlines:
            if outline_curve.get_length() < 0.1:
                continue
            
            rhino_curve = self._internal_geometry_converter.to_rhino_type(outline_curve)
            nurbs_curve = rhino_curve.to_nurbs_curve()
            flat_curves_output.append(nurbs_curve)
        
        results.flat_curves = flat_curves_output
        
        # Process fold line curves
        fold_curves_output = []
        for fold_curve in panel_flat_drawing.fold_lines:
            rhino_curve = self._internal_geometry_converter.to_rhino_type(fold_curve)
            nurbs_curve = rhino_curve.to_nurbs_curve()
            fold_curves_output.append(nurbs_curve)
        
        results.fold_curves = fold_curves_output
        
        # Process relief profile curves
        relief_profiles_output = []
        for relief_profile in panel_flat_drawing.relief_profiles:
            rhino_curve = self._internal_geometry_converter.to_rhino_type(relief_profile)
            nurbs_curve = rhino_curve.to_nurbs_curve()
            relief_profiles_output.append(nurbs_curve)
        
        results.relief_curves = relief_profiles_output
        
        # Process fold labels
        index = 0
        labels_output = []
        for item in panel_flat_drawing.fold_labels:
            # Get corresponding fold line if available
            corresponding_fold = None
            if index < len(panel_flat_drawing.fold_lines):
                corresponding_fold = panel_flat_drawing.fold_lines[index]
            
            # Calculate rotation
            if corresponding_fold is None:
                rotation = 0
            else:
                rotation = corresponding_fold.tangent_at_start.angle_on_plane_to(
                    self._x_axis, 
                    self._z_axis
                )
                if rotation > math.pi / 2:
                    rotation -= math.pi
            
            # Format label text
            text = (f"Index: {item.index_text}, "
                   f"Bend Angle: {item.bend_text}, "
                   f"Fold: {item.orientation_text}")
            
            # Get location for future updates which will return a text dot (for Rhino 8)
            rhino_location = self._internal_geometry_converter.to_rhino_type(item.location)
            
            labels_output.append(text)
            index += 1
        
        results.labels = labels_output
    
    def _create_folded_model_in_rhino(
        self, 
        model_3d, 
        results: SheetMetalFoldingResults
    ) -> None:
        """
        Create 3D folded model and populate results
        
        Args:
            model_3d: IFoldedModel3d object
            results: SheetMetalFoldingResults object to populate
        """
        panel_3d_geometry = model_3d.get_combined_model_3d()
        
        output_meshes = []
        output_breps = []
        
        for solid in panel_3d_geometry:
            rhino_brep = self._internal_geometry_converter.to_rhino_type(solid)
            
            if rhino_brep is None:
                continue
            
            output_breps.append(rhino_brep)
            
            # Create meshes from brep
            meshes = RhinoMesh.create_from_brep(
                rhino_brep, 
                Rhino.Geometry.MeshingParameters.default_analysis_mesh()
            )
            
            # Join and process meshes
            mesh_joined = RhinoMesh()
            mesh_joined.append(meshes)
            mesh_joined.faces.cull_degenerate_faces()
            mesh_joined.vertices.cull_unused()
            mesh_joined.unify_normals()
            mesh_joined.normals.compute_normals()
            
            output_meshes.append(mesh_joined)
        
        results.meshes = output_meshes
        results.breps = output_breps
    
    def _initialize_converters(self, unit_system_manager) -> None:
        """
        Initialize geometry converters (singleton pattern)
        
        Args:
            unit_system_manager: UnitSystemManager instance
        """
        if self._geometry_converter is None:
            RhinoGeometryConverter.initialize(unit_system_manager)
        
        if self._internal_geometry_converter is None:
            InternalGeometryConverter.initialize(unit_system_manager)
        
        self._geometry_converter = RhinoGeometryConverter.instance
        self._internal_geometry_converter = InternalGeometryConverter.instance


# Example usage:
# folder = SheetMetalFolder(
#     mold_model=my_brep,
#     flip=False,
#     material_thickness=0.125,
#     material_k_factor=0.44,
#     bend_radius=0.0625,
#     relief_definition=my_relief_def,
#     input_unit_system=UnitSystem.INCHES
# )
# results = folder.fold()
# print(f"Generated {len(results.flat_curves)} flat curves")
# print(f"Generated {len(results.meshes)} meshes")