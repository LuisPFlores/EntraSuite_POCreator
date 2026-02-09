## EntraSuite_POCreator
Interface to deploy Entra Suite Proof-of-Concept

# Create a new directory for your MCP server
mkdir mcp-entra-server
cd mcp-entra-server

# Create a virtual environment
python -m venv venv

# Activate the virtual environment
# On macOS/Linux:
source venv/bin/activate
# On Windows:
# venv\Scripts\activate

# Install required dependencies
pip install mcp azure-identity msgraph-sdk python-dotenv

