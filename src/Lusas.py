# FEA MCP
# This module provides a Python interface to interact with the LUSAS Modeller using COM.

import win32com.client as win32client
import sys
import logging
#import LPI_22.0 as lpi

logger = logging.getLogger('fea_mcp_server')

class lpi:
    """A class of LUSAS interfaces (no actual use for now)."""
    def IFModeller(self):
        pass
    def IFPoint(self):
        pass
    def IFLine(self):
        pass
    def IFSurface(self):
        pass
    def IFVolume(self):
        pass
    def IFObjectSet(self):
        pass
    def IFGeometryData(self):
        pass

class Lusas:
    def __init__(self, versionString: str = "21.1"):
        self.modeller = None
        self.modeller = self.getModeller(versionString)

    def getModeller(self, versionString: str = "21.1") -> lpi.IFModeller:
        """
        Return Values:
        Modeller (type IFModeller pointer)
        """
        if self.modeller:
            return self.modeller

        # Attach to a running instance of LUSAS
        try:
            # Get the active LUSAS object
            modeller: lpi.IFModeller = win32client.GetActiveObject("Lusas.Modeller." + versionString)
            
        except Exception as e:
            logger.warning(f"No running instance of LUSAS version {versionString} found.")
            return None
        
        logger.info(f"Successfully attached on LUSAS version {versionString}.")
        return modeller
    
# LUSAS server called commands
    def getUnits(self):
        """
        Gets the units of the current model.

        Returns:
        str: A string describing the units of the current model.
        """
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
        try:
            return f"Units of force, length, mass, time and temperature are set to {self.modeller.db().getModelUnits().getName()}"
        except Exception as e:
            logger.error(f"Error getting the model units: {str(e)}")
            return "Error: Failed to get the model units."
        
    def createPoint(self, x:float, y:float, z:float) -> str:
        """
        Creates a point in LUSAS modeller at the specified coordinates.

        Parameters:
        x (float): X coordinate of the point.
        y (float): Y coordinate of the point.
        z (float): Z coordinate of the point.

        Return Values:
        int: ID of the created point.
        """
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
        try:
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setLowerOrderGeometryType("coordinates")
            geom_data.addCoords(x, y, z)
            pnt: lpi.IFPoint = self.modeller.db().createPoint(geom_data).getObjects("Point")[0]
            return f"Point created successfully with ID {pnt.getID()}."
        except Exception as e:
            logger.error(f"Error creating point: {str(e)}")
            return "Error: Failed to create point."

    def createPoints(self, x:list[float], y:list[float], z:list[float]) -> str:
        """
        Creates points in LUSAS modeller at the specified coordinates.

        Parameters:
        x (list): List of x coordinates of the points.
        y (list): List of y coordinates of the points.
        z (list): List of z coordinates of the points.

        Return Values:
        str: IDs of the created point.
        """
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
        try:
            pntIDs = []
            for i in range(len(x)):
                geom_data = self.modeller.geometryData().setAllDefaults()
                geom_data.setLowerOrderGeometryType("coordinates")
                geom_data.addCoords(x[i], y[i], z[i])
                pnt: lpi.IFPoint = self.modeller.db().createPoint(geom_data).getObjects("Point")[0]
                pntIDs.append(pnt.getID())
            return f"Points created successfully with IDs {','.join(map(str, pntIDs))}."
        except Exception as e:
            logger.error(f"Error creating points: {str(e)}")
            return "Error: Failed to create points."

    def createLineByCoordinates(self, x1:float, y1:float, z1:float, x2:float, y2:float, z2:float) -> str:
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
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
        try:
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setCreateMethod("straight")
            geom_data.setLowerOrderGeometryType("coordinates")
            geom_data.addCoords(x1, y1, z1)
            geom_data.addCoords(x2, y2, z2)
            ln : lpi.IFLine = self.modeller.db().createLine(geom_data).getObjects("Line")[0]
            return f"Line created successfully with ID {ln.getID()}."
        except Exception as e:
            logger.error(f"Error creating Line: {str(e)}")
            return "Error: Failed to create line by coordinates."
    
    def createLineByPoints(self, p1: int, p2: int) -> str:
        """
        Creates a line in LUSAS modeller connecting the given points.

        Parameters:
        p1 (int): ID of the first point.
        p2 (int): ID of the second point.

        Return Values:
        int: ID of the created line.
        """
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
        try:
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setCreateMethod("straight")
            geom_data.setLowerOrderGeometryType("points")
            obs = self.modeller.newObjectSet().add("point", p1).add("point", p2)
            ln : lpi.IFLine = obs.createLine(geom_data).getObjects("Line")[0]
            return f"Line created successfully with ID {ln.getID()}."
        except Exception as e:
            logger.error(f"Error creating Line: {str(e)}")
            return "Error: Failed to create line by points."

    def createArcByPoints(self, p1: int, p2: int, p3: int) -> str:
        """
        Creates an arc line in LUSAS modeller connecting the given points.

        Parameters:
        p1 (int): ID of the first point.
        p2 (int): ID of the second point.
        p3 (int): ID of the third point.

        Return Values:
        int: ID of the created line.
        """
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
        try:
            p1 : lpi.IFPoint = self.modeller.db().getObject("Point", p1)
            p2 : lpi.IFPoint = self.modeller.db().getObject("Point", p2)
            p3 : lpi.IFPoint = self.modeller.db().getObject("Point", p3)
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setCreateMethod("arc")
            geom_data.keepMinor()
            geom_data.setStartMiddleEnd()
            geom_data.addCoords(p1.getX(), p1.getY(), p1.getZ())
            geom_data.addCoords(p2.getX(), p2.getY(), p2.getZ())
            geom_data.addCoords(p3.getX(), p3.getY(), p3.getZ())
            geom_data.setLowerOrderGeometryType("coordinates")
            ln : lpi.IFLine = self.modeller.db().createLine(geom_data).getObjects("Line")[0]
            return f"Line created successfully with ID {ln.getID()}."
        except Exception as e:
            logger.error(f"Error creating arc Line: {str(e)}")
            return "Error: Failed to create arc line by points."
        
    def createArcByCoordinates(self, x1:float, y1:float, z1:float, x2:float, y2:float, z2:float, x3:float, y3:float, z3:float) -> str:
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
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
        try:
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setCreateMethod("arc")
            geom_data.keepMinor()
            geom_data.setStartMiddleEnd()
            geom_data.addCoords(x1, y1, z1)
            geom_data.addCoords(x2, y2, z2)
            geom_data.addCoords(x3, y3, z3)
            geom_data.setLowerOrderGeometryType("coordinates")
            ln : lpi.IFLine = self.modeller.db().createLine(geom_data).getObjects("Line")[0]
            return f"Line created successfully with ID {ln.getID()}."
        except Exception as e:
            logger.error(f"Error creating arc Line: {str(e)}")
            return "Error: Failed to create arc line by coordinates."

    def createSplineByPoints(self, pnts:list[int], closeEnds:bool) -> str:
        """
        Creates a spline line in LUSAS modeller connecting the given points.

        Parameters:
        pnts (list): List of point IDs to create the spline.
        closeEnds (bool): Whether to close the ends of the spline.

        Return Values:
        int: ID of the created line.
        """
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
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

            ln : lpi.IFLine = pntsObj.createLine(geom_data).getObjects("Line")[0]
            return f"Line created successfully with ID {ln.getID()}."
        except Exception as e:
            logger.error(f"Error creating spline: {str(e)}")
            return "Error: Failed to create spline by points."

    def createSurfaceByCoordinates(self, x:list[float], y:list[float], z:list[float]) -> str:
        """
        Creates a surface in LUSAS modeller from the given coordinates.

        Parameters:
        x (list): List of x coordinates.
        y (list): List of y coordinates.
        z (list): List of z coordinates.

        Return Values:
        int: ID of the created surface.
        """
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
        try:
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setCreateMethod("coons")
            geom_data.setLowerOrderGeometryType("coordinates")
            for i in range(len(x)):
                geom_data.addCoords(x[i], y[i], z[i])
            surf : lpi.IFSurface = self.modeller.db().createSurface(geom_data).getObjects("Surface")[0]
            return f"Surface created successfully with ID {surf.getID()}."
        except Exception as e:
            logger.error(f"Error creating surface: {str(e)}")
            return "Error: Failed to create surface by coordinates."
    
    def createSurfaceByLines(self, lns:list[int]) -> str:
        """
        Creates a surface in LUSAS modeller from the given lines.

        Parameters:
        lns (list): List of line IDs to create the surface.

        Return Values:
        int: ID of the created surface.
        """
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
        try:
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setCreateMethod("coons")
            geom_data.setLowerOrderGeometryType("lines")
            linesObj = self.modeller.newObjectSet()
            for ln in lns:
                linesObj.add("line", ln)
            surf : lpi.IFSurface = linesObj.createSurface(geom_data).getObjects("Surface")[0]
            return f"Surface created successfully with ID {surf.getID()}."
        except Exception as e:
            logger.error(f"Error creating surface: {str(e)}")
            return "Error: Failed to create surface by lines."

    def createVolume(self, surfs:list[int]) -> str:
        """
        Creates a volume in LUSAS modeller from the given surfaces.
        
        Parameters:
        surfs (list): List of surface IDs to create the volume.

        Return Values:
        int: ID of the created volume.
        """
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
        try:
            geom_data = self.modeller.geometryData().setAllDefaults()
            geom_data.setCreateMethod("solidVolume")
            geom_data.setExtractAllVolumes()
            surfsObj = self.modeller.newObjectSet()
            for surf in surfs:
                surfsObj.add("surface", surf)
            vlm : lpi.IFVolume = self.modeller.db().createVolume(geom_data).getObjects("Volume")[0]
            return f"Volume created successfully with ID {vlm.getID()}."
        except Exception as e:
            logger.error(f"Error creating volume: {str(e)}")
            return "Error: Failed to create volume by surfaces."

    def sweepPoints(self, pnts:list[int], vector: list[float]) -> str:
        """
        Sweeps the given points in the specified direction to create lines.

        Parameters:
        pnts (list[int]): List of point IDs to sweep.
        vector (list[float]): Direction vector for the sweep.

        Return Values:
        list[int]: List of IDs of the created lines.
        """
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
        try:
            myObj = self.modeller.newObjectSet()
            for id in pnts:
                myObj.add("point", id)
            lines : list[lpi.IFLine] = self.sweep_Ext(myObj, vector, "Line").getObjects("Lines")
            return f"Points swept successfully creating lines with IDs {','.join([str(ln.getID()) for ln in lines])}."
        except Exception as e:
            logger.error(f"Error sweeping points: {str(e)}")
            return "Error: Failed to sweep points."

    def sweepLines(self, lines:list[int], vector: list[float]) -> str:
        """
        Sweeps the given lines in the specified direction to create surfaces.

        Parameters:
        lines (list[int]): List of lines IDs to sweep.
        vector (list[float]): Direction vector for the sweep.

        Return Values:
        list[int]: List of IDs of the created surfaces.
        """
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
        try:
            myObj = self.modeller.newObjectSet()
            for id in lines:
                myObj.add("point", id)
            surfs : list[lpi.IFSurface] = self.sweep_Ext(myObj, vector, "Surface").getObjects("Surfaces")
            return f"Lines swept successfully creating surfaces with IDs {','.join([str(surf.getID()) for surf in surfs])}."
        except Exception as e:
            logger.error(f"Error sweeping lines: {str(e)}")
            return "Error: Failed to sweep lines."

    def sweepSurfaces(self, surfs:list[int], vector: list[float]) -> str:
        """
        Sweeps the given surfaces in the specified direction to create volumes.

        Parameters:
        surfs (list[int]): List of surfaces IDs to sweep.
        vector (list[float]): Direction vector for the sweep.

        Return Values:
        list[int]: List of IDs of the created volumes.
        """
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
        try:
            myObj = self.modeller.newObjectSet()
            for id in surfs:
                myObj.add("surface", id)
            vlms : list[lpi.IFVolume] = self.sweep_Ext(myObj, vector, "Volume").getObjects("Volumes")
            return f"Surfaces swept successfully creating volumes with IDs {','.join([str(vlm.getID()) for vlm in vlms])}."
        except Exception as e:
            logger.error(f"Error sweeping surfaces: {str(e)}")
            return "Error: Failed to sweep surfaces."

    def getPoints(self) -> str:
        """Gets all points of the current model."""
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
        try:
            pnts : list[lpi.IFPoint] = self.modeller.db().getObjects("Points")
            data = [f"<Point id={pnt.getID()} x={pnt.getX()}, y={pnt.getY()}, z={pnt.getZ()}>" for pnt in pnts]
            return {' '.join(data)}
        except Exception as e:
            logger.error(f"Error getting all points: {str(e)}")
            return "Error: Failed getting point."
        
    def getLines(self) -> str:
        """Gets all lines of the current model."""
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
        try:
            lns : list[lpi.IFPoint] = self.modeller.db().getObjects("Lines")
            data = [f"<Line id={ln.getID()} x1={ln.getStartPoint().getX()}, y1={ln.getStartPoint().getY()}, z1={ln.getStartPoint().getZ()}, x2={ln.getEndPoint().getX()}, y2={ln.getEndPoint().getY()}, z2={ln.getEndPoint().getZ()}>" for ln in lns]
            return {' '.join(data)}
        except Exception as e:
            logger.error(f"Error getting all lines: {str(e)}")
            return "Error: Failed getting lines."
        
    def select(self, points : list[int], lines : list[int], surfaces : list[int], volumes : list[int]) -> str:
        """
        Selects objects in the current model.
        
        Parameters:
        points (list[int]): List of point IDs to select.
        lines (list[int]): List of line IDs to select.
        surfaces (list[int]): List of surface IDs to select.
        volumes (list[int]): List of volume IDs to select.
        """
        if self.modeller == None:
            return "Error: Not connected to LUSAS."
        
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
    def sweep_Ext(self, trgtObjSet:lpi.IFObjectSet, vector: list[float], hofType:str):
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

    def sweepRot_Ext(self, trgtObjSet:lpi.IFObjectSet, origin:list, hofType:str, degree:float, aboutAxis:str=None):
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
