![fea-mcp-cover](./img/fea-mcp-icon-long.png)

# FEA-MCP Server

An Finite Element Analysis Model Context Protocol Server for AI

## 🚀 Overview

The FEA-MCP Server provides a unified API interface for interacting with various Finite Element Analysis (FEA) software packages. It enables AI control of FEA modelling, analysis, and post-processing through a consistent interface, regardless of the underlying software implementation (currently ETABS and LUSAS are supported).

## ✨ Supported Features

- **Multiple Software**: Supports mainstream FEA software including:
  - [ETABS](https://www.csiamerica.com/products/etabs)
  - [LUSAS](https://www.lusas.com/)
- **Geometric Modelling**:
  - Create Point/Joint
  - Create Line/Frame/Beam/Column
  - Create Volume/Solid
  - Sweep Points/Lines/Surfaces (LUSAS only)
  - Get modelled Points/Lines/Surfaces/Volumes
  - Select objects (LUSAS only)
- **Other**:
  - Read model units

### 🖥️ MCP Tools

The server provides the following main API functions:

- `get_units`: Returns the model units
- `create_objects_by_coordinates`: Batch-creates various geometric objects (points, lines/frames, surfaces/areas, volumes/solids)
- `get_all_geometries`: Returns all the modelled geometric objects (points, lines/frames, surfaces/areas, volumes/solids)
- `get_points`: Returns all the modelled points

    (the following are only available for **ETABS**)

- `get_frames`: Returns all the modelled frames
- `get_areas`: Returns all the modelled areas

    (the following are only available for **LUSAS**)

- `get_lines`: Returns all the modelled lines
- `get_surfaces`: Returns all the modelled surfaces
- `get_volumes`: Returns all the modelled volumes
- `sweep_points`: Sweeps points to create lines
- `sweep_lines`: Sweeps lines to create surfaces
- `sweep_surfaces`: Sweeps surfaces to create volumes
- `select`: Select modelled objects

## 🎯 Future Work

- Model Management: Define materials, sections, loads, and boundary conditions
- Analysis Control: Run simulations and retrieve results
- Coordinate System Support: Work with multiple coordinate systems

## ⚙️ Installation

#### Requirements

Required python libraries:

```
pywin32>=228    # Windows COM interface support
comtypes>=1.4.0 # Windows COM interface support
mcp>=0.1.0      # Model Control Protocol library
```

System Requirements:

- Windows operating system
- Installed FEA software (ETABS, LUSAS)

#### Guide

1. Install the required python libraries from command line:
   
   ```
   pip install pywin32 comtypes mcp
   ```
2. Download this repository and save the extracted files locally (e.g. at ```C:\your_path_to_the_extracted_server\FEA-MCP\```).
3. (Optional) Edit the MCP server configuration file, located at `src/config.json` (see configuration section). By default the server is set to use LUSAS v21.1.
4. Install Claude Desktop (or other AI client with MCP support).
5. Configure Claude Desktop to launch the MCP Server automatically (see Claude Desktop section).
6. You are good to go!

#### Configuration

The configuration file is located at `src/config.json` and contains the following main settings:

```json
{
    "server": {
        "name": "FEA MCP",
        "version": "1.0.0"
    },
    "fea": {
        "software": "LUSAS",
        "version": "21.1"
    }
}
```

- **server**: Server name and version information
- **fea**: 
  - `software`: FEA software (ETABS, LUSAS)
  - `version`: software version (e.g. 21.1 for LUSAS)

## 🤖 AI Clients

#### 5ire

Open 5ire > Tools > New, input the following info and then click Save:

| Input       | Value                                                               |
| ----------- | ------------------------------------------------------------------- |
| Tool Key    | *fea*                                                               |
| Description | *Finite Elements Analysis connection server (ETABS, LUSAS)*         |
| Command     | `python C:\your_path_to_the_extracted_server\FEA-MCP\src\server.py` |

**Caution**: update the path! (single slashes)

Then turn on the server and you are good to go!

#### Claude Desktop

Open Claude Desktop and navigate to `File > Settings > Developer > Edit Config`, edit `claude_desktop_config.json` and add the following JSON.

```json
{
    "mcpServers": {
        "fea": {
            "command": "python",
            "args": [
                "C:\\your_path_to_the_extracted_server\\FEA-MCP\\src\\server.py"
            ]
        }
    }
}
```

Caution: update the path and use double backslash!
Then restart Claude Desktop (from the tray icon, right click > Quit).