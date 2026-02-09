# EntraSuite_POCreator
Interface to deploy Entra Suite Proof-of-Concept

## Step 1: Set Up Your Python Environment
### Create a new directory for your MCP server
```
mkdir mcp-entra-server
cd mcp-entra-server
```
### Create a virtual environment
```
python -m venv venv
```
### Activate the virtual environment
### On macOS/Linux:
```
source venv/bin/activate
```
### On Windows:
```
venv\Scripts\activate
```
### Install required dependencies
```
   pip install mcp azure-identity msgraph-sdk python-dotenv
```

## Step 2: Update Your Entra App Registration
### API Permissions needed:

User.Read.All - Read all users' full profiles

Group.Read.All - Read all groups

Directory.Read.All - Read directory data

User.ReadWrite.All - Read and write all users (if you want write access)

Group.ReadWrite.All - Read and write all groups (if you want write access)

**Make sure to Grant admin consent for these permissions.**

## Step 3: Create Environment Configuration
Create a .env file:

```
AZURE_CLIENT_ID=your_client_id_here
AZURE_TENANT_ID=your_tenant_id_here
AZURE_CLIENT_SECRET=your_client_secret_here  # Optional, for app-only auth
```
