# FEA MCP
# This module provides a Python interface to interact with the LUSAS Modeller using COM.

import sys
import logging
import win32com.client as win32client
from mcp.server.fastmcp import Context
from pydantic import BaseModel

logger = logging.getLogger('fea_mcp_server')

# Help Classes for data definition
class GeomObject(BaseModel):
    """A class representing a point/line/surface/volume to be created by points."""
    type : str
    """Type of the geometry (can be "point", "straight line", "arc", "spline", "surface", "volume")."""
    xs : list[float]
    """List of x coordinates."""
    ys : list[float]
    """List of y coordinates."""
    zs : list[float]
    """List of z coordinates."""
    id : int = -1
    """Object ID (-1 if not created yet)."""
    selected : bool = False
    """Whether the object is selected or not."""

class Lusas:
    def __init__(self, versionString: str = "21.1"):
        self.versionString = versionString
        self.modeller : 'IFModeller' = None
        self.set_modeller(False)
    
    def set_modeller(self, createModel : bool = True) -> bool:
        """Checks LUSAS connection and returns True if connected."""
        if self.modeller:
            try:
                if not self.modeller.existsDatabase() and createModel:
                    self.modeller.newProject()
                return True
            except Exception as e:
                logger.warning(f"Invalid modeller reference ({str(e)}). Trying to reconnect.")
                self.modeller = None

        # Attach to a running instance of LUSAS
        try:
            # Get the active LUSAS object
            self.modeller: 'IFModeller' = win32client.GetActiveObject("Lusas.Modeller." + self.versionString)
            
        except Exception as e:
            logger.warning(f"No running instance of LUSAS version {self.versionString} found.")
            return False
        
        logger.info(f"Successfully attached on LUSAS version {self.versionString}.")

        if not self.modeller.existsDatabase() and createModel:
            self.modeller.newProject()
        return True

# LUSAS server called commands
    def get_units(self):
        """
        Gets the units of the current model.

        Returns:
        str: A string describing the units of the current model.
        """
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            return f"Units of force, length, mass, time and temperature are set to {self.modeller.db().getModelUnits().getName()}"
        except Exception as e:
            logger.error(f"Error getting the model units: {str(e)}")
            return "Error: Failed to get the model units."
        
    def create_point(self, x:float, y:float, z:float) -> str:
        """
        Creates a point in LUSAS modeller at the specified coordinates.

        Parameters:
        x (float): X coordinate of the point.
        y (float): Y coordinate of the point.
        z (float): Z coordinate of the point.

        Return Values:
        int: ID of the created point.
        """
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setLowerOrderGeometryType("coordinates")
            geom_data.addCoords(x, y, z)
            pnt: 'IFPoint' = self.modeller.db().createPoint(geom_data).getObjects("Point")[0]
            return f"Point created successfully with ID {pnt.getID()}."
        except Exception as e:
            logger.error(f"Error creating point: {str(e)}")
            return "Error: Failed to create point."

    def create_points(self, x:list[float], y:list[float], z:list[float]) -> str:
        """
        Creates points in LUSAS modeller at the specified coordinates.

        Parameters:
        x (list): List of x coordinates of the points.
        y (list): List of y coordinates of the points.
        z (list): List of z coordinates of the points.

        Return Values:
        str: IDs of the created point.
        """
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            pntIDs = []
            for i in range(len(x)):
                geom_data = self.modeller.geometryData().setAllDefaults()
                geom_data.setLowerOrderGeometryType("coordinates")
                geom_data.addCoords(x[i], y[i], z[i])
                pnt: 'IFPoint' = self.modeller.db().createPoint(geom_data).getObjects("Point")[0]
                pntIDs.append(pnt.getID())
            return f"Points created successfully with IDs {','.join(map(str, pntIDs))}."
        except Exception as e:
            logger.error(f"Error creating points: {str(e)}")
            return "Error: Failed to create points."

    def create_line_by_coordinates(self, x1:float, y1:float, z1:float, x2:float, y2:float, z2:float) -> str:
        """
        Creates a line in LUSAS modeller connecting the given points.

        Parameters:
        x1 (float): X coordinate of the start point.
        y1 (float): Y coordinate of the start point.
        z1 (float): Z coordinate of the start point.
        x2 (float): X coordinate of the end point.
        y2 (float): Y coordinate of the end point.
        z2 (float): Z coordinate of the end point.

        Return Values:
        int: ID of the created line.
        """
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setCreateMethod("straight")
            geom_data.setLowerOrderGeometryType("coordinates")
            geom_data.addCoords(x1, y1, z1)
            geom_data.addCoords(x2, y2, z2)
            ln : 'IFLine' = self.modeller.db().createLine(geom_data).getObjects("Line")[0]
            return f"Line created successfully with ID {ln.getID()}."
        except Exception as e:
            logger.error(f"Error creating Line: {str(e)}")
            return "Error: Failed to create line by coordinates."
    
    def create_line_by_points(self, p1: int, p2: int) -> str:
        """
        Creates a line in LUSAS modeller connecting the given points.

        Parameters:
        p1 (int): ID of the first point.
        p2 (int): ID of the second point.

        Return Values:
        int: ID of the created line.
        """
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setCreateMethod("straight")
            geom_data.setLowerOrderGeometryType("points")
            obs = self.modeller.newObjectSet().add("point", p1).add("point", p2)
            ln : 'IFLine' = obs.createLine(geom_data).getObjects("Line")[0]
            return f"Line created successfully with ID {ln.getID()}."
        except Exception as e:
            logger.error(f"Error creating Line: {str(e)}")
            return "Error: Failed to create line by points."

    def create_arc_by_points(self, p1: int, p2: int, p3: int) -> str:
        """
        Creates an arc line in LUSAS modeller connecting the given points.

        Parameters:
        p1 (int): ID of the first point.
        p2 (int): ID of the second point.
        p3 (int): ID of the third point.

        Return Values:
        int: ID of the created line.
        """
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            p1 : 'IFPoint' = self.modeller.db().getObject("Point", p1)
            p2 : 'IFPoint' = self.modeller.db().getObject("Point", p2)
            p3 : 'IFPoint' = self.modeller.db().getObject("Point", p3)
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setCreateMethod("arc")
            geom_data.keepMinor()
            geom_data.setStartMiddleEnd()
            geom_data.addCoords(p1.getX(), p1.getY(), p1.getZ())
            geom_data.addCoords(p2.getX(), p2.getY(), p2.getZ())
            geom_data.addCoords(p3.getX(), p3.getY(), p3.getZ())
            geom_data.setLowerOrderGeometryType("coordinates")
            ln : 'IFLine' = self.modeller.db().createLine(geom_data).getObjects("Line")[0]
            return f"Line created successfully with ID {ln.getID()}."
        except Exception as e:
            logger.error(f"Error creating arc Line: {str(e)}")
            return "Error: Failed to create arc line by points."
        
    def create_arc_by_coordinates(self, x1:float, y1:float, z1:float, x2:float, y2:float, z2:float, x3:float, y3:float, z3:float) -> str:
        """
        Creates an arc line in LUSAS modeller connecting the given points.

        Parameters:
        x1 (float): X coordinate of the first point.
        y1 (float): Y coordinate of the first point.
        z1 (float): Z coordinate of the first point.
        x2 (float): X coordinate of the second point.
        y2 (float): Y coordinate of the second point.
        z2 (float): Z coordinate of the second point.
        x3 (float): X coordinate of the third point.
        y3 (float): Y coordinate of the third point.
        z3 (float): Z coordinate of the third point.

        Return Values:
        int: ID of the created line.
        """
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setCreateMethod("arc")
            geom_data.keepMinor()
            geom_data.setStartMiddleEnd()
            geom_data.addCoords(x1, y1, z1)
            geom_data.addCoords(x2, y2, z2)
            geom_data.addCoords(x3, y3, z3)
            geom_data.setLowerOrderGeometryType("coordinates")
            ln : 'IFLine' = self.modeller.db().createLine(geom_data).getObjects("Line")[0]
            return f"Line created successfully with ID {ln.getID()}."
        except Exception as e:
            logger.error(f"Error creating arc Line: {str(e)}")
            return "Error: Failed to create arc line by coordinates."

    def create_spline_by_coordinates(self, x:list[float], y:list[float], z:list[float], closeEnds:bool) -> str:
        """
        Creates a spline line in LUSAS modeller connecting the given points.

        Parameters:
        x (list): List of x coordinates.
        y (list): List of y coordinates.
        z (list): List of z coordinates.
        closeEnds (bool): Whether to close the ends of the spline.

        Return Values:
        int: ID of the created line.
        """
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setCreateMethod("spline") #coons
            geom_data.useSelectionOrder(True)
            if closeEnds:
                geom_data.closeEndPoints(True)
            geom_data.setLowerOrderGeometryType("coordinates")
            for i in range(len(x)):
                geom_data.addCoords(x[i], y[i], z[i])
            ln : 'IFLine' = self.modeller.db().createLine(geom_data).getObjects("Line")[0]
            return f"Line created successfully with ID {ln.getID()}."
        except Exception as e:
            logger.error(f"Error creating spline: {str(e)}")
            return "Error: Failed to create spline by points."
        
    def create_spline_by_points(self, pnts:list[int], closeEnds:bool) -> str:
        """
        Creates a spline line in LUSAS modeller connecting the given points.

        Parameters:
        pnts (list): List of point IDs to create the spline.
        closeEnds (bool): Whether to close the ends of the spline.

        Return Values:
        int: ID of the created line.
        """
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setCreateMethod("spline")
            geom_data.useSelectionOrder(True)
            if closeEnds:
                geom_data.closeEndPoints(True)
            geom_data.setLowerOrderGeometryType("points")

            pntsObj = self.modeller.newObjectSet()
            for pnt in pnts:
                pntsObj.add("point", pnt)

            ln : 'IFLine' = pntsObj.createLine(geom_data).getObjects("Line")[0]
            return f"Line created successfully with ID {ln.getID()}."
        except Exception as e:
            logger.error(f"Error creating spline: {str(e)}")
            return "Error: Failed to create spline by points."

    def create_surface_by_coordinates(self, x:list[float], y:list[float], z:list[float]) -> str:
        """
        Creates a surface in LUSAS modeller from the given coordinates.

        Parameters:
        x (list): List of x coordinates.
        y (list): List of y coordinates.
        z (list): List of z coordinates.

        Return Values:
        int: ID of the created surface.
        """
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setCreateMethod("coons")
            geom_data.setLowerOrderGeometryType("coordinates")
            for i in range(len(x)):
                geom_data.addCoords(x[i], y[i], z[i])
            surf : 'IFSurface' = self.modeller.db().createSurface(geom_data).getObjects("Surface")[0]
            return f"Surface created successfully with ID {surf.getID()}."
        except Exception as e:
            logger.error(f"Error creating surface: {str(e)}")
            return "Error: Failed to create surface by coordinates."
    
    def create_surface_by_lines(self, lns:list[int]) -> str:
        """
        Creates a surface in LUSAS modeller from the given lines.

        Parameters:
        lns (list): List of line IDs to create the surface.

        Return Values:
        int: ID of the created surface.
        """
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setCreateMethod("coons")
            geom_data.setLowerOrderGeometryType("lines")
            linesObj = self.modeller.newObjectSet()
            for ln in lns:
                linesObj.add("line", ln)
            surf : 'IFSurface' = linesObj.createSurface(geom_data).getObjects("Surface")[0]
            return f"Surface created successfully with ID {surf.getID()}."
        except Exception as e:
            logger.error(f"Error creating surface: {str(e)}")
            return "Error: Failed to create surface by lines."

    def create_volume(self, surfs:list[int]) -> str:
        """
        Creates a volume in LUSAS modeller from the given surfaces.
        
        Parameters:
        surfs (list): List of surface IDs to create the volume.

        Return Values:
        int: ID of the created volume.
        """
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setCreateMethod("solidVolume")
            geom_data.setExtractAllVolumes()
            surfsObj = self.modeller.newObjectSet()
            for surf in surfs:
                surfsObj.add("surface", surf)
            vlm : 'IFVolume' = self.modeller.db().createVolume(geom_data).getObjects("Volume")[0]
            return f"Volume created successfully with ID {vlm.getID()}."
        except Exception as e:
            logger.error(f"Error creating volume: {str(e)}")
            return "Error: Failed to create volume by surfaces."

    async def create_objects_by_coordinates(self, objects: list[GeomObject], ctx: Context) -> list[str]:
        """
        Batch-creates various geometric objects in LUSAS modeller. All objects can be created in one call.
        This is useful for creating multiple objects at once, such as points, lines, arcs, and surfaces.
        When creating an object, there is no need to specify the lower order geometry type, as the are automatically created.

        Parameters:
        objects (list of GeomObject): A list of object definitions.
            Each object must include a "type" key indicating the object type, and relevant parameters (coordinates).

        Returns:
        list of str: Status messages for each object created.
        """
        #- "spline": {"points": list[int], "closeEnds": bool}
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        results = []
        try:
            self.modeller.db().beginCommandBatch("MCP create objects", True)
            for i in range(len(objects)):
                # Report progress to the client
                await ctx.report_progress(i, len(objects))
                # Create object
                obj : GeomObject = objects[i]
                try:
                    obj_type = obj.type.lower()
                    if obj_type == "point":
                        results.append(self.create_point(obj.xs[0], obj.ys[0], obj.zs[0]))
                    elif obj_type == "straight line":
                        results.append(self.create_line_by_coordinates(obj.xs[0], obj.ys[0], obj.zs[0],
                                                                    obj.xs[1], obj.ys[1], obj.zs[1]))
                    elif obj_type == "arc":
                        results.append(self.create_arc_by_coordinates(obj.xs[0], obj.ys[0], obj.zs[0],
                                                                obj.xs[1], obj.ys[1], obj.zs[1],
                                                                obj.xs[2], obj.ys[2], obj.zs[2]))
                    elif obj_type == "spline":
                        results.append(self.create_spline_by_coordinates(obj.xs, obj.ys, obj.zs, False))
                    elif obj_type == "surface":
                        results.append(self.create_surface_by_coordinates(obj.xs, obj.ys, obj.zs))
                    else:
                        results.append(f"Error: Unsupported type '{obj_type}'.")
                except Exception as e:
                    results.append(f"Error processing {obj_type}: {str(e)}")
        except Exception as e:
            return f"Error creation objects: {str(e)}"
        finally:
            self.modeller.db().closeCommandBatch()
            # Fit model in view
            self.modeller.view().scaleToFit()
        return results


    def sweep_points(self, pnts:list[int], vector: list[float]) -> str:
        """
        Sweeps the given points in the specified direction to create lines.

        Parameters:
        pnts (list[int]): List of point IDs to sweep.
        vector (list[float]): Direction vector for the sweep.

        Return Values:
        list[int]: List of IDs of the created lines.
        """
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            myObj = self.modeller.newObjectSet()
            for id in pnts:
                myObj.add("point", id)
            lines : list['IFLine'] = self.sweep_Ext(myObj, vector, "Line").getObjects("Lines")
            return f"Points swept successfully creating lines with IDs {','.join([str(ln.getID()) for ln in lines])}."
        except Exception as e:
            logger.error(f"Error sweeping points: {str(e)}")
            return "Error: Failed to sweep points."

    def sweep_lines(self, lines:list[int], vector: list[float]) -> str:
        """
        Sweeps the given lines in the specified direction to create surfaces.

        Parameters:
        lines (list[int]): List of lines IDs to sweep.
        vector (list[float]): Direction vector for the sweep.

        Return Values:
        list[int]: List of IDs of the created surfaces.
        """
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            myObj = self.modeller.newObjectSet()
            for id in lines:
                myObj.add("point", id)
            surfs : list['IFSurface'] = self.sweep_Ext(myObj, vector, "Surface").getObjects("Surfaces")
            return f"Lines swept successfully creating surfaces with IDs {','.join([str(surf.getID()) for surf in surfs])}."
        except Exception as e:
            logger.error(f"Error sweeping lines: {str(e)}")
            return "Error: Failed to sweep lines."

    def sweep_surfaces(self, surfs:list[int], vector: list[float]) -> str:
        """
        Sweeps the given surfaces in the specified direction to create volumes.

        Parameters:
        surfs (list[int]): List of surfaces IDs to sweep.
        vector (list[float]): Direction vector for the sweep.

        Return Values:
        list[int]: List of IDs of the created volumes.
        """
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            myObj = self.modeller.newObjectSet()
            for id in surfs:
                myObj.add("surface", id)
            vlms : list['IFVolume'] = self.sweep_Ext(myObj, vector, "Volume").getObjects("Volumes")
            return f"Surfaces swept successfully creating volumes with IDs {','.join([str(vlm.getID()) for vlm in vlms])}."
        except Exception as e:
            logger.error(f"Error sweeping surfaces: {str(e)}")
            return "Error: Failed to sweep surfaces."

    # This is a bit slow
    async def get_geometries(self, ctx: Context) -> list[GeomObject] | str:
        """Gets all geometries (points, lines, surfaces or volumes) of the current model."""
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            geoms : list[GeomObject] = []

            await ctx.report_progress(0, 4)
            points : list['IFPoint'] = self.modeller.db().getObjects("Points")
            for pnt in points:
                geoms.append(GeomObject(type="point", xs=[pnt.getX()], ys=[pnt.getY()], zs=[pnt.getZ()], id=pnt.getID(), selected=pnt.isSelected()))

            await ctx.report_progress(1, 4)
            lines : list['IFLine'] = self.modeller.db().getObjects("Lines")
            for line in lines:
                #TODO: Check if arc
                #if line.getTypeCode()
                l_points : list['IFPoint'] = self.modeller.newObjectSet().add(line).addLOF("points").getObjects("Points")
                geoms.append(GeomObject(type="line", xs=[p.getX() for p in l_points], ys=[p.getY() for p in l_points], zs=[p.getZ() for p in l_points], id=line.getID(), selected=line.isSelected()))
                
            await ctx.report_progress(2, 4)
            surfaces : list['IFSurface'] = self.modeller.db().getObjects("Surfaces")
            for surface in surfaces:
                l_points : list['IFPoint'] = self.modeller.newObjectSet().add(surface).addLOF("points").getObjects("Points")
                geoms.append(GeomObject(type="surface", xs=[p.getX() for p in l_points], ys=[p.getY() for p in l_points], zs=[p.getZ() for p in l_points], id=surface.getID(), selected=surface.isSelected()))
                
            await ctx.report_progress(3, 4)
            volumes : list['IFVolume'] = self.modeller.db().getObjects("Volumes")
            for volume in volumes:
                l_points : list['IFPoint'] = self.modeller.newObjectSet().add(volume).addLOF("points").getObjects("Points")
                geoms.append(GeomObject(type="volume", xs=[p.getX() for p in l_points], ys=[p.getY() for p in l_points], zs=[p.getZ() for p in l_points], id=volume.getID(), selected=volume.isSelected()))
                
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
            points : list['IFPoint'] = self.modeller.db().getObjects("Points")
            for pnt in points:
                geoms.append(GeomObject(type="point", xs=[pnt.getX()], ys=[pnt.getY()], zs=[pnt.getZ()]))
            
            #data = [f"<Point id={pnt.getID()} x={pnt.getX()}, y={pnt.getY()}, z={pnt.getZ()}>" for pnt in pnts]
            return geoms
        except Exception as e:
            logger.error(f"Error getting all points: {str(e)}")
            return "Error: Failed getting point."
        
    def get_lines(self) -> list[GeomObject] | str:
        """Gets all lines of the current model."""
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            geoms : list[GeomObject] = []
            lines : list['IFLine'] = self.modeller.db().getObjects("Lines")
            for line in lines:
                #TODO: Check if arc
                #if line.getTypeCode()
                l_points : list['IFPoint'] = self.modeller.newObjectSet().add(line).addLOF("points").getObjects("Points")
                geoms.append(GeomObject(type="line", xs=[p.getX() for p in l_points], ys=[p.getY() for p in l_points], zs=[p.getZ() for p in l_points], id=line.getID(), selected=line.isSelected()))

            #data = [f"<Line id={ln.getID()} x1={ln.getStartPoint().getX()}, y1={ln.getStartPoint().getY()}, z1={ln.getStartPoint().getZ()}, x2={ln.getEndPoint().getX()}, y2={ln.getEndPoint().getY()}, z2={ln.getEndPoint().getZ()}>" for ln in lns]
            return geoms
        except Exception as e:
            logger.error(f"Error getting all lines: {str(e)}")
            return "Error: Failed getting lines."
        
    def get_surfaces(self) -> list[GeomObject] | str:
        """Gets all surfaces of the current model."""
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            geoms : list[GeomObject] = []
            surfaces : list['IFSurface'] = self.modeller.db().getObjects("Surfaces")
            for surface in surfaces:
                l_points : list['IFPoint'] = self.modeller.newObjectSet().add(surface).addLOF("points").getObjects("Points")
                geoms.append(GeomObject(type="surface", xs=[p.getX() for p in l_points], ys=[p.getY() for p in l_points], zs=[p.getZ() for p in l_points], id=surface.getID(), selected=surface.isSelected()))
                
            return geoms
        except Exception as e:
            logger.error(f"Error getting all surfaces: {str(e)}")
            return "Error: Failed getting surfaces."
        
    def get_volumes(self) -> list[GeomObject] | str:
        """Gets all volumes of the current model."""
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            geoms : list[GeomObject] = []
            volumes : list['IFVolume'] = self.modeller.db().getObjects("Volumes")
            for volume in volumes:
                l_points : list['IFPoint'] = self.modeller.newObjectSet().add(volume).addLOF("points").getObjects("Points")
                geoms.append(GeomObject(type="volume", xs=[p.getX() for p in l_points], ys=[p.getY() for p in l_points], zs=[p.getZ() for p in l_points], id=volume.getID(), selected=volume.isSelected()))
                
            return geoms
        except Exception as e:
            logger.error(f"Error getting all volumes: {str(e)}")
            return "Error: Failed getting volumes."
        
    def select(self, points : list[int], lines : list[int], surfaces : list[int], volumes : list[int]) -> str:
        """
        Selects objects in the current model.
        
        Parameters:
        points (list[int]): List of point IDs to select.
        lines (list[int]): List of line IDs to select.
        surfaces (list[int]): List of surface IDs to select.
        volumes (list[int]): List of volume IDs to select.
        """
        if not self.set_modeller():
            return f"Error: Cannot connect on LUSAS version {self.versionString}."
        
        try:
            self.modeller.selection().remove("all")
            for id in points:
                self.modeller.selection().add("point", id)
            for id in lines:
                self.modeller.selection().add("line", id)
            for id in surfaces:
                self.modeller.selection().add("surface", id)
            for id in volumes:
                self.modeller.selection().add("volume", id)
            return "Objects selected."
        
        except Exception as e:
            logger.error(f"Error selecting model objects: {str(e)}")
            return "Error: Failed to select objects."

# LUSAS extensions (not called directly from the server)
    def sweep_Ext(self, trgtObjSet:'IFObjectSet', vector: list[float], hofType:str):
        types = ["Point", "Line", "Surface", "Volume"]
        MaximumDimension = types.index(hofType)

        attr = self.modeller.db().createTranslationTransAttr("Temp_SweepTranslation", vector)
        attr.setSweepType("straight")
        attr.setHofType(hofType)

        geomData = self.modeller.newGeometryData()
        geomData.setMaximumDimension(MaximumDimension)
        geomData.setTransformation(attr)
        geomData.sweptArcType("straight")

        objSet = trgtObjSet.sweep(geomData)
        self.modeller.db().deleteAttribute(attr)

        return objSet

    def sweepRot_Ext(self, trgtObjSet:'IFObjectSet', origin:list, hofType:str, degree:float, aboutAxis:str=None):
        types = ["Point", "Line", "Surface", "Volume"]
        MaximumDimension = types.index(hofType)

        if aboutAxis is None:
            aboutAxis = "z"

        title = "Temp_SweepRotation"
        if aboutAxis.lower() == "x":
            attr = self.modeller.db().createYZRotationTransAttr(title, degree, origin)
        elif aboutAxis.lower() == "y":
            attr = self.modeller.db().createXZRotationTransAttr(title, degree, origin)
        else:
            attr = self.modeller.db().createXYRotationTransAttr(title, degree, origin)

        attr.setSweepType("minorArc")
        attr.setHofType(hofType)

        geomData = self.modeller.newGeometryData()
        geomData.setMaximumDimension(MaximumDimension)
        geomData.setTransformation(attr)
        geomData.sweptArcType("minorArc")

        objSet = trgtObjSet.sweep(geomData)
        self.modeller.db().deleteAttribute(attr)

        return objSet


# This is for testing purposes only
if __name__ == "__main__":
    modeller = Lusas().modeller
    if modeller is None:
        sys.exit("Failed to connect to LUSAS.")
        
    if not modeller.existsDatabase():
        print("Database not found, creating a project...")
        modeller.newProject()
