# FEA MCP
# Config loader

import json
import logging
import os

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger('fea_mcp')

class Config:

    def __init__(self):
        # Load config
        configPath = os.path.join(os.path.dirname(__file__), 'config.json')
        try:
            with open(configPath, 'r', encoding='utf-8') as f:
                self.data = json.load(f)
            logger.info("Configuration loaded successfully.")
        except Exception as e:
            logger.error(f"Could not load config file: {str(e)}")
            # Return default config if loading fails
            self.data = {
                "server": {
                    "name": "FEA MCP",
                    "version": "1.0.0"
                },
                "fea": {
                    "software": "LUSAS",
                    "version": "21.1"
                }
            }

    @property
    def serverName(self) -> str:
        return self.data['server']['name']

    @property
    def serverVersion(self) -> str:
        return self.data['server']['version']
    
    @property
    def feaName(self) -> str:
        return self.data['fea']['software'].upper()

    @property
    def feaVersion(self) -> str:
        version = float(self.data['fea']['version'])
        return f"{version:.1f}"

# This is for testing purposes only
if __name__ == "__main__":
    config = Config()
    print(f"Server Name: {config.serverName}")
    print(f"Server Version: {config.serverVersion}")
    print(f"FEA Software: {config.feaName}")
    print(f"FEA Version: {config.feaVersion}")
