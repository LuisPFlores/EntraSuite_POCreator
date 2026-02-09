# EntraSuite_POCreator
This interface enables deployment of Entra Suite configurations and creation of Proof of Concepts (PoCs) using Claude AI Platform with a local MCP server

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

*User.Read.All* - Read all users' full profiles

*Group.Read.All* - Read all groups

*Directory.Read.All* - Read directory data

*User.ReadWrite.All* - Read and write all users (if you want write access)

*Group.ReadWrite.All* - Read and write all groups (if you want write access)

*Policy.Read.All* - Read conditional access policies

*Policy.ReadWrite.ConditionalAccess* - Manage conditional access

*Application.Read.All* - Read applications

*Application.ReadWrite.All* - Manage applications (if needed)

*NetworkAccess.ReadWrite.All* - Manage network access policies

*NetworkAccessPolicy.ReadWrite.All* - Read/Write network access policies

*Policy.Read.All* - Read policies

*Policy.ReadWrite.ConditionalAccess* - Manage conditional access

*AccessReview.ReadWrite.All* - Manage access reviews

*EntitlementManagement.ReadWrite.All* - Manage entitlement management

*PrivilegedAccess.ReadWrite.AzureAD* - Manage PIM

*LifecycleWorkflows.ReadWrite.All* - Manage lifecycle workflows

*RoleManagement.ReadWrite.Directory* - Manage directory roles

**To provide the client app access to Microsoft Graph:** https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-configure-app-access-web-apis  

*Make sure to Grant admin consent for these permissions.*

## Step 3: Create Environment Configuration
Create a .env file:

```
AZURE_CLIENT_ID=your_client_id_here
AZURE_TENANT_ID=your_tenant_id_here
AZURE_CLIENT_SECRET=your_client_secret_here  # Optional, for app-only auth
```

## Step 4: Configure Claude Desktop
Edit Claude Desktop config file:
```
macOS: ~/Library/Application Support/Claude/claude_desktop_config.json
```
```
Windows: %APPDATA%\Claude\claude_desktop_config.json
```
Add this configuration:
```
{
  "mcpServers": {
    "entra": {
      "command": "/full/path/to/mcp-entra-server/venv/bin/python",
      "args": ["/full/path/to/mcp-entra-server/server.py"],
      "env": {
        "AZURE_CLIENT_ID": "your_client_id",
        "AZURE_TENANT_ID": "your_tenant_id",
        "AZURE_CLIENT_SECRET": "your_client_secret"
      }
    }
  }
```
}

Replace the paths and credentials with your actual values.
## Step 5: Restart Claude Desktop
Close and reopen Claude Desktop completely.

# Try asking me:

## Entra Private Access Operations:

* "List all private access applications"

* "Create a new Quick Access app for our internal CRM at crm.company.local"

* "Assign the Sales team group to the CRM application"

* "Show me all private access connectors"

## Combined Operations:

* "Create a forwarding profile for private access"

* "Set up a private access app for our internal SharePoint server at sharepoint.internal.company.com on port 443"

* "Assign users admin@tenant.onmicrosoft.com and name@tenant.onmicrosoft.com to the internal app"

## Access Reviews:

* "Create an access review for the IT Department group with a 14-day duration"

* "List all active access reviews"

* "Show me the decisions from the latest access review"

## Entitlement Management:

* "List all access packages in our tenant"

* "Create a new catalog called 'Sales Resources'"

* "Request access to the 'SharePoint Sales' package for user name@tenant.onmicrosoft.com"

## Privileged Identity Management (PIM):

* "List all privileged roles available"

* "Show me the role assignments for name@tenant.onmicrosoft.com"

* "Create an eligible assignment for Global Administrator role"

* "Activate the Security Administrator role with justification 'Security incident investigation'"

## Lifecycle Workflows:

* "List all lifecycle workflows"

* "Create a new joiner workflow for onboarding"

* "Run the onboarding workflow for the new user name@tenant.onmicrosoft.com"
