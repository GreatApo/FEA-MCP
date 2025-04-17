# FEA MCP
# This module provides a Python interface to interact with ETABS using COM.

import sys
import logging
import comtypes.client
from mcp.server.fastmcp import Context
from pydantic import BaseModel

logger = logging.getLogger('fea_mcp_server')

# Help Classes for data definition
class GeomObject(BaseModel):
    """A class representing a point/line/surface to be created by points."""
    type : str
    """Type of the geometry (can be "point", "line", "surface")."""
    xs : list[float]
    """List of x coordinates."""
    ys : list[float]
    """List of y coordinates."""
    zs : list[float]
    """List of z coordinates."""
    id : str = ""
    """Object ID (empty if not created yet)."""

class Etabs:
    def __init__(self):
        self.SapModel = None
        self.set_modeller(False)

    def set_modeller(self, createModel : bool = True) -> bool:
        """
        Connect to a running instance of ETABS.
        If no instance is found, a new instance will be created.

        Return Values:
        SapModel (type cOAPI pointer)
        """
        if self.SapModel:
            try:
                if False and createModel: # TODO: Check if file is open
                    self.SapModel.InitializeNewModel(6) # Set units to kN, m, C
                    self.SapModel.File.NewBlank()
                return True
            except Exception as e:
                logger.warning(f"Invalid modeller reference ({str(e)}). Trying to reconnect.")
                self.SapModel = None
        
        # Attach to a running instance of ETABS
        comtypes.CoInitialize()
        try:
            # Create API helper object
            helper = comtypes.client.CreateObject('ETABSv1.Helper')
            helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper)
            EtabsObject = helper.GetObject("CSI.ETABS.API.ETABSObject")

            if not EtabsObject:
                logger.info("No running instance of the program found. Trying to lunch a new instance.")
                # Create a new instance of ETABS
                EtabsObject = helper.GetObject("CSI.ETABS.API.ETABSObject")
                EtabsObject.ApplicationStart()
                if createModel:
                    EtabsObject.SapModel.InitializeNewModel(6)  # Set units to kN, m, C
                    EtabsObject.SapModel.File.NewBlank()

            logger.info(f"ETABS connected. OAPI Version Number: {EtabsObject.GetOAPIVersionNumber()}")

        except Exception as e:
            logger.error("No running instance of the program found or failed to attach.")
            return False
        
        finally:
            comtypes.CoUninitialize()
        
        self.SapModel = EtabsObject.SapModel
        return True

    def get_version(self) -> str:
        """Return model version"""
        version, myVersionNumber, ret = self.SapModel.GetVersion()
        if ret != 0:
            return f"Error getting version"
        return version

    def get_units(self):
        """
        Gets the units of the current model.

        Returns:
        str: A string describing the units of the current model.
        """
        if not self.set_modeller():
            return "Error: Cannot connect on ETABS."
        
        presetUnits = ["lb, in, F", "lb, ft, F", "kip, in, F", "kip, ft, F", "kN, mm, C", "kN, m, C", "kgf, mm, C", "kgf, m, C", "N, mm, C", "N, m, C", "Ton, mm, C", "Ton, m, C", "kN, cm, C", "kgf, cm, C", "N, cm, C", "Ton, cm, C"]
        MyUnits = self.SapModel.GetPresentUnits()
        if MyUnits < 1 or MyUnits > len(presetUnits):
            return "Unknown units"
        return f"Units of force, length and temperature are set to {presetUnits[MyUnits-1]}"

    def save(self):
        """Saves the current model."""
        if not self.set_modeller():
            return "Error: Cannot connect on ETABS."
        
        ret = self.SapModel.File.Save()
        if ret != 0:
            return "Error saving the model"
        return f"Model saved successfully."

    def create_joint(self, x: float, y: float, z: float) -> str:
        """Creates a point/joint.

        Args:
            SapModel: SapModel object
            x: X coordinate
            y: Y coordinate
            z: Z coordinate
        """
        if not self.set_modeller():
            return "Error: Cannot connect on ETABS."
        
        pName, ret = self.SapModel.PointObj.AddCartesian(x, y, z)
        if ret != 0:
            return f"Error adding joint ({x}, {y}, {z}) to the model."
        ret = self.SapModel.View.RefreshView()
        return f"Joint created successfully with ID {pName}."

    def create_frame(self, xi: float, yi: float, zi: float, xj: float, yj: float, zj: float) -> str:
        """Creates a line/frame.

        Args:
            SapModel: SapModel object
            xi: Start X coordinate
            yi: Start Y coordinate
            zi: Start Z coordinate
            xj: End X coordinate
            yj: End Y coordinate
            zj: End Z coordinate
        """
        if not self.set_modeller():
            return "Error: Cannot connect on ETABS."
        
        fName, ret = self.SapModel.FrameObj.AddByCoord(xi, yi, zi, xj, yj, zj)
        if ret != 0:
            return f"Error adding line/frame ({xi}, {yi}, {zi}) - ({xj}, {yj}, {zj}) to the model."
        #ret = self.SapModel.FrameObj.SetLocalAxes(fName, 0)
        #if ret != 0:
        #    return f"Error adding frame rotation to the model."
        ret = self.SapModel.View.RefreshView()
        return f"Frame created successfully with ID {fName}."

    def create_area(self, x:list[float], y:list[float], z:list[float]) -> int:
        """
        Creates a surface in ETABS modeller from the given coordinates.

        Parameters:
        x (list): List of x coordinates.
        y (list): List of y coordinates.
        z (list): List of z coordinates.

        Return Values:
        int: ID of the created surface.
        """
        if not self.set_modeller():
            return "Error: Cannot connect on ETABS."
        
        x, y, z, aName, ret = self.SapModel.AreaObj.AddByCoord(len(x), x, y, z)
        if ret != 0:
            return f"Error adding surface/area to the model."
        ret = self.SapModel.View.RefreshView()
        return f"Area created successfully with ID {aName}."

    def create_solid(self, x:list[float], y:list[float], z:list[float]) -> int:
        """
        Creates a solid in ETABS modeller from the given coordinates.

        Parameters:
        x (list): List of x coordinates.
        y (list): List of y coordinates.
        z (list): List of z coordinates.

        Return Values:
        int: ID of the created solid.
        """
        if not self.set_modeller():
            return "Error: Cannot connect on ETABS."
        
        # Not working
        x, y, z, sName, ret = self.SapModel.SolidObj.AddByCoord(x, y, z)
        if ret != 0:
            return f"Error adding volume/solid to the model."
        ret = self.SapModel.View.RefreshView()
        return f"Solid created successfully with ID {sName}."

    def create_objects_by_coordinates(self, objects: list[GeomObject], ctx: Context) -> list[str]:
        """
        Batch-creates various geometric objects in ETABS modeller. All objects can be created in one call.
        This is useful for creating multiple objects at once, such as points, lines and surfaces.
        When creating an object, there is no need to specify the lower order geometry type, as the are automatically created.

        Parameters:
        objects (list of GeomObject): A list of object definitions.
            Each object must include a "type" key indicating the object type, and relevant parameters (coordinates).

        Returns:
        list of str: Status messages for each object created.
        """
        #- "spline": {"points": list[int], "closeEnds": bool}
        if not self.set_modeller():
            return f"Error: Cannot connect on ETABS version {self.versionString}."
        
        results = []
        try:
            for i in range(len(objects)):
                # Report progress to the client
                ctx.report_progress(i, len(objects))
                # Create object
                obj : GeomObject = objects[i]
                try:
                    obj_type = obj.type.lower()
                    if obj_type == "point":
                        results.append(self.create_joint(obj.xs[0], obj.ys[0], obj.zs[0]))
                    elif obj_type == "line":
                        results.append(self.create_frame(obj.xs[0], obj.ys[0], obj.zs[0],
                                                                    obj.xs[1], obj.ys[1], obj.zs[1]))
                    elif obj_type == "surface":
                        results.append(self.create_area(obj.xs, obj.ys, obj.zs))
                    else:
                        results.append(f"Error: Unsupported type '{obj_type}'.")
                except Exception as e:
                    results.append(f"Error processing {obj_type}: {str(e)}")
        except Exception as e:
            return f"Error creation objects: {str(e)}"
        return results

    # This is a bit slow
    async def get_geometries(self, ctx: Context) -> list[GeomObject] | str:
        """Gets all geometries (points, lines, surfaces or volumes) of the current model."""
        if not self.set_modeller():
            return f"Error: Cannot connect on ETABS version {self.versionString}."
        
        try:
            geoms : list[GeomObject] = []
            
            # Get all points
            await ctx.report_progress(0, 3)
            [numberPts, ptNames, ptX, ptY, ptZ, ptCsys] = self.SapModel.PointObj.GetAllPoints()
            for i in range(numberPts):
                geoms.append(GeomObject(type="point", xs=[ptX[i]], ys=[ptY[i]], zs=[ptZ[i]], id=ptNames[i]))

            # Get all lines/frames
            await ctx.report_progress(1, 3)
            frame_objs = self.SapModel.FrameObj.GetAllFrames()
            for i in range(frame_objs[0]):
                frameNm = frame_objs[1][i]
                #prop = frame_objs[2][i]
                #story = frame_objs[3][i]
                #pt1 = frame_objs[4][i]
                #pt2 = frame_objs[5][i]
                x1 = frame_objs[6][i]
                y1 = frame_objs[7][i]
                z1 = frame_objs[8][i]
                x2 = frame_objs[9][i]
                y2 = frame_objs[10][i]
                z2 = frame_objs[11][i]
                geoms.append(GeomObject(type="line", xs=[x1,x2], ys=[y1,y2], zs=[z1,z2], id=frameNm))
            
            # Get all areas/surfaces
            await ctx.report_progress(2, 3)
            (n, names, design, _, delim, _, x_coords, y_coords, z_coords, _) = self.SapModel.AreaObj.GetAllAreas()
            i = 0
            for count, j in enumerate(delim):
                name = names[count]
                xs = x_coords[i: j + 1]
                ys = y_coords[i: j + 1]
                zs = z_coords[i: j + 1]
                geoms.append(GeomObject(type="surface", xs=xs, ys=ys, zs=zs, id=name))
                i = j + 1
            
            logger.info(f"Get geometries:")
            logger.info(geoms)

            if len(geoms) == 0:
                return "No geometries found in the model."
            return geoms
        except Exception as e:
            logger.error(f"Error getting geometries: {str(e)}")
            return "Error: Failed getting geometries."
        
    def get_points(self) -> list[GeomObject] | str:
        """Gets all points of the current model."""
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            geoms : list[GeomObject] = []
            [numberPts, ptNames, ptX, ptY, ptZ, ptCsys] = self.SapModel.PointObj.GetAllPoints()
            for i in range(numberPts):
                geoms.append(GeomObject(type="point", xs=[ptX[i]], ys=[ptY[i]], zs=[ptZ[i]], id=ptNames[i]))
            return geoms
        except Exception as e:
            logger.error(f"Error getting all points: {str(e)}")
            return "Error: Failed getting point."
        
    def get_frames(self) -> list[GeomObject] | str:
        """Gets all frames/lines of the current model."""
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            geoms : list[GeomObject] = []
            frame_objs = self.SapModel.FrameObj.GetAllFrames()
            for i in range(frame_objs[0]):
                frameNm = frame_objs[1][i]
                #prop = frame_objs[2][i]
                #story = frame_objs[3][i]
                #pt1 = frame_objs[4][i]
                #pt2 = frame_objs[5][i]
                x1 = frame_objs[6][i]
                y1 = frame_objs[7][i]
                z1 = frame_objs[8][i]
                x2 = frame_objs[9][i]
                y2 = frame_objs[10][i]
                z2 = frame_objs[11][i]
                geoms.append(GeomObject(type="line", xs=[x1,x2], ys=[y1,y2], zs=[z1,z2], id=frameNm))
            return geoms
        except Exception as e:
            logger.error(f"Error getting all frames: {str(e)}")
            return "Error: Failed getting frames."
        
    def get_areas(self) -> list[GeomObject] | str:
        """Gets all areas/surfaces of the current model."""
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            geoms : list[GeomObject] = []
            (n, names, design, _, delim, _, x_coords, y_coords, z_coords, _) = self.SapModel.AreaObj.GetAllAreas()
            i = 0
            for count, j in enumerate(delim):
                name = names[count]
                xs = x_coords[i: j + 1]
                ys = y_coords[i: j + 1]
                zs = z_coords[i: j + 1]
                geoms.append(GeomObject(type="surface", xs=xs, ys=ys, zs=zs, id=name))
                i = j + 1
            return geoms
        except Exception as e:
            logger.error(f"Error getting all areas: {str(e)}")
            return "Error: Failed getting areas."
        
# This is for testing purposes only
if __name__ == "__main__":
    modeller = Etabs()
    if modeller.SapModel is None:
        sys.exit("Failed to connect to ETABS.")

    model = modeller.SapModel
    print(f"Model File path: {model.GetModelFilename()}")
    print(f"Version: {modeller.get_version()}")
    print(modeller.get_units())

    msg = modeller.create_joint(0, 0, 0)
    print(msg)
    msg = modeller.create_frame(0, 0, 0, 2, 1, 0)
    print(msg)

    x = [0, 1, 1, 0]
    y = [0, 0, 1, 1]
    z = [0, 0, 0, 0]
    msg = modeller.create_area(x, y, z)
    print(msg)

    x = [0, 1, 1, 0, 0, 1, 1, 0]
    y = [0, 0, 1, 1, 0, 0, 1, 1]
    z = [0, 0, 0, 0, 1, 1, 1, 1]
    # Sap2000 only:
    #msg = modeller.createSolid(x, y, z)
    #print(msg)
