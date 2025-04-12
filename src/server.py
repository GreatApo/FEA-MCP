# FEA MCP
# The main server module for the FEA MCP (Model Creation Protocol) server.
# This module handles the connection to the FEA software, manages the model, and provides the available interface commands.

from mcp.server.fastmcp import FastMCP
import logging
from config import *
from Etabs import *
from Lusas import *

 # Constants
supportedSoftware = ["LUSAS", "ETABS"]

# Initialize FastMCP server
mcp = FastMCP("fea")

# Start logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('fea_mcp.log', encoding='utf-8')
    ]
)

logger = logging.getLogger('fea_mcp_server')
logger.info("Starting FEA MCP server")

# Load configuration
config = Config()
logger.info("Set software is %s", config.feaName)

def getSoftwareObj() -> object:
    try:
        if config.feaName == "LUSAS":
            modeller = Lusas(config.feaVersion)
            
        elif config.feaName == "ETABS":
            modeller = Etabs()
        
        #elif config.feaName == "SAP2000":

    except Exception as e:
        logger.error(f"Failed to connect on {config.feaName} instance with error: {str(e)}")
        return None
    
    return modeller

@mcp.tool()
def connectOnSoftware() -> str:
    """Connects on the target Finite Elements software modeller."""
    global feaSoftware
    if feaSoftware != None:
        return f"Already connected on {config.feaName} version {config.feaVersion}."
    
    if config.feaName not in supportedSoftware:
        return f"Unsupported software. Please set from: {', '.join(supportedSoftware)}."

    # Update software object
    feaSoftware = getSoftwareObj()
    if not feaSoftware:
        return f"Failed to connect to the specified software instance ({config.feaName}, {config.feaVersion})."
    return f"Connected on {config.feaName} version {config.feaVersion}"

@mcp.tool()
def getSoftware() -> str:
    """Gets the target Finite Elements software name and version."""
    return f"The target FEA software is {config.feaName} version {config.feaVersion}"

@mcp.tool()
def getModelUnits() -> str:
    """
    Gets the units of the current model.

    Returns:
    str: A string describing the units of the current model.
    """

    if not feaSoftware:
        return "Open and connect a software instance first."
    
    return feaSoftware.getUnits()

@mcp.tool()
def createPoint(x: float, y: float, z: float) -> str:
    """Creates a point in the FEA software.

    Args:
        x: X coordinate
        y: Y coordinate
        z: Z coordinate
    """
    
    if not feaSoftware:
        return "Open and connect a software instance first."
    
    if config.feaName == "ETABS" or config.feaName == "SAP2000":
        id : int = feaSoftware.createJoint(x, y, z)
        
    elif config.feaName == "LUSAS":
        return feaSoftware.createPoint(x, y, z)

    if id.startswith("Error"):
        return id
    return f"Point created successfully with ID {id}."

@mcp.tool()
def createLine(x1:float, y1:float, z1:float, x2:float, y2:float, z2:float) -> str:
    """
    Creates a line (aka beam/frame/column) in the FEA software.

    Parameters:
    x1 (float): X coordinate of the starting point.
    y1 (float): Y coordinate of the starting point.
    z1 (float): Z coordinate of the starting point.
    x2 (float): X coordinate of the ending point.
    y2 (float): Y coordinate of the ending point.
    z2 (float): Z coordinate of the ending point.

    Return Values:
    int: ID of the created line.
    """
    
    if not feaSoftware:
        return "Open and connect a software instance first."
    
    if config.feaName == "ETABS" or config.feaName == "SAP2000":
        id : str = feaSoftware.createFrame(x1, y1, z1, x2, y2, z2)
        
    elif config.feaName == "LUSAS":
        return feaSoftware.createLineByCoordinates(x1, y1, z1, x2, y2, z2)
    
    if id.startswith("Error"):
        return id

    return f"Line created successfully with ID {id}."

@mcp.tool()
def createSurface(x:list[float], y:list[float], z:list[float]) -> str:
    """
    Creates a surface (aka area/slab) in the FEA software.

    Parameters:
    x (list): List of x coordinates.
    y (list): List of y coordinates.
    z (list): List of z coordinates.

    Return Values:
    int: ID of the created surface.
    """
    
    if not feaSoftware:
        return "Open and connect a software instance first."
    
    if config.feaName == "ETABS" or config.feaName == "SAP2000":
        id : str = feaSoftware.createArea(x, y, z)
        
    elif config.feaName == "LUSAS":
        return feaSoftware.createSurfaceByCoordinates(x, y, z)
    
    if id.startswith("Error"):
        return id

    return f"Line created successfully with ID {id}."

# Connect to the software
feaSoftware = getSoftwareObj()

# Register software specific tools (for commands available only on specific software, or not implemented yet)
if config.feaName == "LUSAS":
    lusas : Lusas = feaSoftware

    logger.info(f"Registering LUSAS specific tools...")
    createPoints = mcp.tool()(lusas.createPoints)
    createLineByPoints = mcp.tool()(lusas.createLineByPoints)
    createArcByPoints = mcp.tool()(lusas.createArcByPoints)
    createArcByCoordinates = mcp.tool()(lusas.createArcByCoordinates)
    createSurfaceByLines = mcp.tool()(lusas.createSurfaceByLines)
    sweepPoints = mcp.tool()(lusas.sweepPoints)
    sweepLines = mcp.tool()(lusas.sweepLines)
    sweepSurfaces = mcp.tool()(lusas.sweepSurfaces)

elif config.feaName == "ETABS":
    # No specific tools to registere yet
    pass

if __name__ == "__main__":
    # Initialize and run the server
    mcp.run(transport='stdio')
    