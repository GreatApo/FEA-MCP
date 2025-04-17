# FEA MCP
# The main server module for the FEA MCP (Model Creation Protocol) server.
# This module handles the connection to the FEA software, manages the model, and provides the available interface commands.

from mcp.server.fastmcp import FastMCP, Context
import logging
from config import *
from Etabs import *
from Lusas import *

 # Constants
supportedSoftware = ["LUSAS", "ETABS"]


# Initialize FastMCP server
mcp = FastMCP("fea", dependencies=["pywin32", "comtypes"])

# Start logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('fea_mcp.log', encoding='utf-8')
    ])
logger = logging.getLogger('fea_mcp_server')
logger.info("Starting FEA MCP server")

# Load configuration
config = Config()
logger.info("Set software is %s", config.feaName)


@mcp.resource("config://app")
def get_config() -> str:
    """Static server configuration data"""
    return config.data

# Register software specific tools (for commands available only on specific software, or not implemented yet)
if config.feaName == "LUSAS":
    lusas = Lusas(config.feaVersion)
    logger.info(f"Registering LUSAS tools...")

    get_units = mcp.tool()(lusas.get_units)

    # Geometry creation tools
    #create_points = mcp.tool()(lusas.create_points)
    #create_line_by_points = mcp.tool()(lusas.create_line_by_points)
    #create_arc_by_points = mcp.tool()(lusas.create_arc_by_points)
    #create_arc_by_coordinates = mcp.tool()(lusas.create_arc_by_coordinates)
    #create_surface_by_lines = mcp.tool()(lusas.create_surface_by_lines)
    create_objects_by_coordinates = mcp.tool()(lusas.create_objects_by_coordinates)

    # Geometry operations tools
    sweep_points = mcp.tool()(lusas.sweep_points)
    sweep_lines = mcp.tool()(lusas.sweep_lines)
    sweep_surfaces = mcp.tool()(lusas.sweep_surfaces)
    
    # Pull Geometry tools
    get_all_geometries = mcp.tool()(lusas.get_geometries) # a bit slow
    get_points = mcp.tool()(lusas.get_points)
    get_lines = mcp.tool()(lusas.get_lines)
    get_surfaces = mcp.tool()(lusas.get_surfaces)
    get_volumes = mcp.tool()(lusas.get_volumes)

    select = mcp.tool()(lusas.select)

elif config.feaName == "ETABS":
    etabs = Etabs()
    logger.info(f"Registering ETABS tools...")

    get_units = mcp.tool()(etabs.get_units)
    #save = mcp.tool()(etabs.save)

    # Geometry creation tools
    #create_joint = mcp.tool()(etabs.create_joint)
    #create_frame = mcp.tool()(etabs.create_frame)
    #create_area = mcp.tool()(etabs.create_area)
    create_objects_by_coordinates = mcp.tool()(etabs.create_objects_by_coordinates)
    
    # Pull Geometry tools
    get_all_geometries = mcp.tool()(etabs.get_geometries) # a bit slow
    get_points = mcp.tool()(etabs.get_points)
    get_frames = mcp.tool()(etabs.get_frames)
    get_areas = mcp.tool()(etabs.get_areas)

if __name__ == "__main__":
    # Initialize and run the server
    mcp.run(transport='stdio')
    