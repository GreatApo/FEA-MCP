# FEA MCP
# This module provides a Python interface to interact with ETABS using COM.

import comtypes.client
import win32com.client as win32client
import sys
import logging

logger = logging.getLogger('fea_mcp_server')

class Etabs:
    def __init__(self):
        self.SapModel = self.connect()

    def connect(self):
        """
        Connect to a running instance of ETABS.
        If no instance is found, a new instance will be created.

        Return Values:
        SapModel (type cOAPI pointer)
        """
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
                EtabsObject.SapModel.InitializeNewModel(6)  # Set units to kN, m, C
                EtabsObject.SapModel.File.NewBlank()

            logger.info(f"ETABS connected. OAPI Version Number: {EtabsObject.GetOAPIVersionNumber()}")

        except Exception as e:
            logger.error("No running instance of the program found or failed to attach.")
            return None
        
        finally:
            comtypes.CoUninitialize()
        
        return EtabsObject.SapModel

    def getVersion(self) -> str:
        """Return model version"""
        version, myVersionNumber, ret = self.SapModel.GetVersion()
        if ret != 0:
            return f"Error getting version"
        return version

    def getUnits(self):
        """
        Gets the units of the current model.

        Returns:
        str: A string describing the units of the current model.
        """
        presetUnits = ["lb, in, F", "lb, ft, F", "kip, in, F", "kip, ft, F", "kN, mm, C", "kN, m, C", "kgf, mm, C", "kgf, m, C", "N, mm, C", "N, m, C", "Ton, mm, C", "Ton, m, C", "kN, cm, C", "kgf, cm, C", "N, cm, C", "Ton, cm, C"]
        MyUnits = self.SapModel.GetPresentUnits()
        if MyUnits < 1 or MyUnits > len(presetUnits):
            return "Unknown units"
        return f"Units of force, length and temperature are set to {presetUnits[MyUnits-1]}"

    def save(self):
        """Saves the current model."""
        ret = self.SapModel.File.Save()
        if ret != 0:
            return "Error saving the model"
        return f"Model saved successfully."

    def createJoint(self, x: float, y: float, z: float) -> str:
        """Creates a point/joint.

        Args:
            SapModel: SapModel object
            x: X coordinate
            y: Y coordinate
            z: Z coordinate
        """
        pName, ret = self.SapModel.PointObj.AddCartesian(x, y, z)
        if ret != 0:
            return f"Error adding point ({x}, {y}, {z}) to the model."
        ret = self.SapModel.View.RefreshView()
        return pName

    def createFrame(self, xi: float, yi: float, zi: float, xj: float, yj: float, zj: float) -> str:
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
        fName, ret = self.SapModel.FrameObj.AddByCoord(xi, yi, zi, xj, yj, zj)
        if ret != 0:
            return f"Error adding line/frame ({xi}, {yi}, {zi}) - ({xj}, {yj}, {zj}) to the model."
        #ret = self.SapModel.FrameObj.SetLocalAxes(fName, 0)
        #if ret != 0:
        #    return f"Error adding frame rotation to the model."
        ret = self.SapModel.View.RefreshView()
        return fName

    def createArea(self, x:list[float], y:list[float], z:list[float]) -> int:
        """
        Creates a surface in ETABS modeller from the given coordinates.

        Parameters:
        x (list): List of x coordinates.
        y (list): List of y coordinates.
        z (list): List of z coordinates.

        Return Values:
        int: ID of the created surface.
        """
        x, y, z, aName, ret = self.SapModel.AreaObj.AddByCoord(len(x), x, y, z)
        if ret != 0:
            return f"Error adding surface/area to the model."
        ret = self.SapModel.View.RefreshView()
        return aName

    def createSolid(self, x:list[float], y:list[float], z:list[float]) -> int:
        """
        Creates a solid in ETABS modeller from the given coordinates.

        Parameters:
        x (list): List of x coordinates.
        y (list): List of y coordinates.
        z (list): List of z coordinates.

        Return Values:
        int: ID of the created solid.
        """
        # Not working
        x, y, z, sName, ret = self.SapModel.SolidObj.AddByCoord(x, y, z)
        if ret != 0:
            return f"Error adding volume/solid to the model."
        ret = self.SapModel.View.RefreshView()
        return sName

# This is for testing purposes only
if __name__ == "__main__":
    modeller = Etabs()
    if modeller.SapModel is None:
        sys.exit("Failed to connect to ETABS.")

    model = modeller.SapModel
    print(f"Model File path: {model.GetModelFilename()}")
    print(f"Version: {modeller.getVersion()}")
    print(modeller.getUnits())

    msg = modeller.createJoint(0, 0, 0)
    print(msg)
    msg = modeller.createFrame(0, 0, 0, 2, 1, 0)
    print(msg)

    x = [0, 1, 1, 0]
    y = [0, 0, 1, 1]
    z = [0, 0, 0, 0]
    msg = modeller.createArea(x, y, z)
    print(msg)

    x = [0, 1, 1, 0, 0, 1, 1, 0]
    y = [0, 0, 1, 1, 0, 0, 1, 1]
    z = [0, 0, 0, 0, 1, 1, 1, 1]
    msg = modeller.createSolid(x, y, z)
    print(msg)
