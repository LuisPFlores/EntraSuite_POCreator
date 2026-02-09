#!/usr/bin/env python3

import asyncio
import os
import json
from typing import Any
from datetime import datetime, timedelta
from dotenv import load_dotenv
import aiohttp

from mcp.server.models import InitializationOptions
from mcp.server import NotificationOptions, Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

from azure.identity import ClientSecretCredential, DeviceCodeCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.users_request_builder import UsersRequestBuilder
from msgraph.generated.groups.groups_request_builder import GroupsRequestBuilder

# Load environment variables
load_dotenv()

# Configuration
CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
TENANT_ID = os.getenv('AZURE_TENANT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')

# Global Graph client and access token
graph_client = None
access_token = None


def get_credential():
    """Get appropriate credential based on available configuration"""
    if CLIENT_SECRET:
        return ClientSecretCredential(
            tenant_id=TENANT_ID,
            client_id=CLIENT_ID,
            client_secret=CLIENT_SECRET
        )
    else:
        return DeviceCodeCredential(
            tenant_id=TENANT_ID,
            client_id=CLIENT_ID
        )


async def get_access_token():
    """Get access token for Graph API calls"""
    global access_token
    
    if access_token is None:
        credential = get_credential()
        token = credential.get_token('https://graph.microsoft.com/.default')
        access_token = token.token
    
    return access_token


async def initialize_graph_client():
    """Initialize the Microsoft Graph client"""
    global graph_client
    
    if graph_client is None:
        credential = get_credential()
        scopes = ['https://graph.microsoft.com/.default']
        graph_client = GraphServiceClient(credentials=credential, scopes=scopes)
    
    return graph_client


async def make_graph_request(method: str, endpoint: str, data: dict = None):
    """Make a Graph API request using aiohttp for endpoints not in SDK"""
    token = await get_access_token()
    
    url = f"https://graph.microsoft.com/beta/{endpoint}"
    
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    
    async with aiohttp.ClientSession() as session:
        if method.upper() == 'GET':
            async with session.get(url, headers=headers) as response:
                if response.status >= 400:
                    error_text = await response.text()
                    raise Exception(f"API Error {response.status}: {error_text}")
                return await response.json()
        
        elif method.upper() == 'POST':
            async with session.post(url, headers=headers, json=data) as response:
                if response.status >= 400:
                    error_text = await response.text()
                    raise Exception(f"API Error {response.status}: {error_text}")
                return await response.json()
        
        elif method.upper() == 'PATCH':
            async with session.patch(url, headers=headers, json=data) as response:
                if response.status >= 400:
                    error_text = await response.text()
                    raise Exception(f"API Error {response.status}: {error_text}")
                return await response.json()
        
        elif method.upper() == 'DELETE':
            async with session.delete(url, headers=headers) as response:
                if response.status >= 400:
                    error_text = await response.text()
                    raise Exception(f"API Error {response.status}: {error_text}")
                return {"status": "deleted"}


# Create server instance
server = Server("mcp-entra-complete")


@server.list_tools()
async def handle_list_tools() -> list[Tool]:
    """List all available Entra tools (ID, SSE, Governance)"""
    return [
        # ===== ENTRA ID USER/GROUP MANAGEMENT =====
        Tool(
            name="list_users",
            description="List users from Microsoft Entra ID with optional filtering",
            inputSchema={
                "type": "object",
                "properties": {
                    "top": {"type": "number", "description": "Number of users to retrieve (default: 10, max: 999)"},
                    "filter": {"type": "string", "description": "OData filter query (e.g., \"startswith(displayName,'John')\")"},
                    "search": {"type": "string", "description": "Search query for users"}
                },
            },
        ),
        Tool(
            name="get_user",
            description="Get detailed information about a specific user by ID or UPN",
            inputSchema={
                "type": "object",
                "properties": {"user_id": {"type": "string", "description": "User ID (object ID) or User Principal Name (email)"}},
                "required": ["user_id"],
            },
        ),
        Tool(
            name="list_groups",
            description="List groups from Microsoft Entra ID",
            inputSchema={
                "type": "object",
                "properties": {
                    "top": {"type": "number", "description": "Number of groups to retrieve (default: 10)"},
                    "filter": {"type": "string", "description": "OData filter query"}
                },
            },
        ),
        Tool(
            name="get_group",
            description="Get detailed information about a specific group",
            inputSchema={
                "type": "object",
                "properties": {"group_id": {"type": "string", "description": "Group ID (object ID)"}},
                "required": ["group_id"],
            },
        ),
        Tool(
            name="get_group_members",
            description="Get members of a specific group",
            inputSchema={
                "type": "object",
                "properties": {
                    "group_id": {"type": "string", "description": "Group ID (object ID)"},
                    "top": {"type": "number", "description": "Number of members to retrieve (default: 10)"}
                },
                "required": ["group_id"],
            },
        ),
        Tool(
            name="get_user_groups",
            description="Get groups that a user is a member of",
            inputSchema={
                "type": "object",
                "properties": {"user_id": {"type": "string", "description": "User ID (object ID) or User Principal Name"}},
                "required": ["user_id"],
            },
        ),
        Tool(
            name="search_users",
            description="Search for users across multiple fields (displayName, mail, userPrincipalName)",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "Search query"},
                    "top": {"type": "number", "description": "Number of results (default: 10)"}
                },
                "required": ["query"],
            },
        ),
        Tool(
            name="get_user_manager",
            description="Get the manager of a specific user",
            inputSchema={
                "type": "object",
                "properties": {"user_id": {"type": "string", "description": "User ID (object ID) or User Principal Name"}},
                "required": ["user_id"],
            },
        ),
        Tool(
            name="get_user_direct_reports",
            description="Get direct reports of a specific user",
            inputSchema={
                "type": "object",
                "properties": {"user_id": {"type": "string", "description": "User ID (object ID) or User Principal Name"}},
                "required": ["user_id"],
            },
        ),
        
        # ===== INTERNET ACCESS MANAGEMENT =====
        Tool(
            name="get_internet_access_status",
            description="Get the status of Entra Internet Access (Global Secure Access)",
            inputSchema={"type": "object", "properties": {}},
        ),
        Tool(
            name="list_forwarding_profiles",
            description="List all traffic forwarding profiles",
            inputSchema={"type": "object", "properties": {}},
        ),
        Tool(
            name="get_forwarding_profile",
            description="Get details of a specific forwarding profile",
            inputSchema={
                "type": "object",
                "properties": {"profile_id": {"type": "string", "description": "Forwarding profile ID"}},
                "required": ["profile_id"],
            },
        ),
        Tool(
            name="create_forwarding_profile",
            description="Create a new traffic forwarding profile",
            inputSchema={
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Name of the forwarding profile"},
                    "description": {"type": "string", "description": "Description of the profile"},
                    "profile_type": {"type": "string", "description": "Type: 'microsoftTraffic' or 'privateAccess' or 'internetAccess'", "enum": ["microsoftTraffic", "privateAccess", "internetAccess"]},
                    "priority": {"type": "number", "description": "Priority (lower number = higher priority)"}
                },
                "required": ["name", "profile_type"],
            },
        ),
        Tool(
            name="update_forwarding_profile",
            description="Update an existing forwarding profile",
            inputSchema={
                "type": "object",
                "properties": {
                    "profile_id": {"type": "string", "description": "Forwarding profile ID"},
                    "name": {"type": "string", "description": "New name"},
                    "description": {"type": "string", "description": "New description"},
                    "state": {"type": "string", "description": "enabled or disabled", "enum": ["enabled", "disabled"]}
                },
                "required": ["profile_id"],
            },
        ),
        Tool(
            name="delete_forwarding_profile",
            description="Delete a forwarding profile",
            inputSchema={
                "type": "object",
                "properties": {"profile_id": {"type": "string", "description": "Forwarding profile ID"}},
                "required": ["profile_id"],
            },
        ),
        
        # ===== SECURITY POLICIES =====
        Tool(
            name="list_security_policies",
            description="List all web content filtering policies",
            inputSchema={"type": "object", "properties": {}},
        ),
        Tool(
            name="get_security_policy",
            description="Get details of a specific security policy",
            inputSchema={
                "type": "object",
                "properties": {"policy_id": {"type": "string", "description": "Security policy ID"}},
                "required": ["policy_id"],
            },
        ),
        Tool(
            name="create_security_policy",
            description="Create a new web content filtering security policy",
            inputSchema={
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Policy name"},
                    "description": {"type": "string", "description": "Policy description"},
                    "priority": {"type": "number", "description": "Priority (lower = higher priority)"},
                    "action": {"type": "string", "description": "Action to take: 'allow' or 'block'", "enum": ["allow", "block"]}
                },
                "required": ["name", "action"],
            },
        ),
        Tool(
            name="add_fqdn_to_policy",
            description="Add a Fully Qualified Domain Name (FQDN) to a security policy",
            inputSchema={
                "type": "object",
                "properties": {
                    "policy_id": {"type": "string", "description": "Security policy ID"},
                    "fqdn": {"type": "string", "description": "Domain name (e.g., 'example.com' or '*.example.com')"}
                },
                "required": ["policy_id", "fqdn"],
            },
        ),
        Tool(
            name="add_web_category_to_policy",
            description="Add a web category to a security policy",
            inputSchema={
                "type": "object",
                "properties": {
                    "policy_id": {"type": "string", "description": "Security policy ID"},
                    "category": {"type": "string", "description": "Category name (e.g., 'Gambling', 'AdultContent', 'SocialNetworking')"}
                },
                "required": ["policy_id", "category"],
            },
        ),
        Tool(
            name="update_security_policy",
            description="Update an existing security policy",
            inputSchema={
                "type": "object",
                "properties": {
                    "policy_id": {"type": "string", "description": "Security policy ID"},
                    "name": {"type": "string", "description": "New name"},
                    "description": {"type": "string", "description": "New description"},
                    "state": {"type": "string", "description": "enabled or disabled", "enum": ["enabled", "disabled"]}
                },
                "required": ["policy_id"],
            },
        ),
        Tool(
            name="delete_security_policy",
            description="Delete a security policy",
            inputSchema={
                "type": "object",
                "properties": {"policy_id": {"type": "string", "description": "Security policy ID"}},
                "required": ["policy_id"],
            },
        ),
        
        # ===== PRIVATE ACCESS =====
        Tool(
            name="list_private_access_apps",
            description="List all Quick Access applications (Private Access)",
            inputSchema={"type": "object", "properties": {}},
        ),
        Tool(
            name="get_private_access_app",
            description="Get details of a specific Quick Access application",
            inputSchema={
                "type": "object",
                "properties": {"app_id": {"type": "string", "description": "Application ID"}},
                "required": ["app_id"],
            },
        ),
        Tool(
            name="create_private_access_app",
            description="Create a new Quick Access application for Private Access",
            inputSchema={
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Application name"},
                    "description": {"type": "string", "description": "Application description"},
                    "fqdn": {"type": "string", "description": "Fully qualified domain name (e.g., 'app.company.com')"},
                    "ip_addresses": {"type": "array", "description": "Array of IP addresses or ranges", "items": {"type": "string"}},
                    "ports": {"type": "array", "description": "Array of ports (e.g., [80, 443])", "items": {"type": "number"}}
                },
                "required": ["name"],
            },
        ),
        Tool(
            name="assign_users_to_private_access_app",
            description="Assign users or groups to a Quick Access application",
            inputSchema={
                "type": "object",
                "properties": {
                    "app_id": {"type": "string", "description": "Application ID"},
                    "user_ids": {"type": "array", "description": "Array of user IDs", "items": {"type": "string"}},
                    "group_ids": {"type": "array", "description": "Array of group IDs", "items": {"type": "string"}}
                },
                "required": ["app_id"],
            },
        ),
        Tool(
            name="update_private_access_app",
            description="Update an existing Quick Access application",
            inputSchema={
                "type": "object",
                "properties": {
                    "app_id": {"type": "string", "description": "Application ID"},
                    "name": {"type": "string", "description": "New name"},
                    "description": {"type": "string", "description": "New description"},
                    "state": {"type": "string", "description": "enabled or disabled", "enum": ["enabled", "disabled"]}
                },
                "required": ["app_id"],
            },
        ),
        Tool(
            name="delete_private_access_app",
            description="Delete a Quick Access application",
            inputSchema={
                "type": "object",
                "properties": {"app_id": {"type": "string", "description": "Application ID"}},
                "required": ["app_id"],
            },
        ),
        Tool(
            name="list_private_access_connectors",
            description="List all Global Secure Access connectors for Private Access",
            inputSchema={"type": "object", "properties": {}},
        ),
        Tool(
            name="get_private_access_connector",
            description="Get details of a specific connector",
            inputSchema={
                "type": "object",
                "properties": {"connector_id": {"type": "string", "description": "Connector ID"}},
                "required": ["connector_id"],
            },
        ),
        
        # ===== TRAFFIC ANALYSIS =====
        Tool(
            name="get_traffic_logs",
            description="Get recent traffic logs from Entra Internet Access",
            inputSchema={
                "type": "object",
                "properties": {"top": {"type": "number", "description": "Number of logs to retrieve (default: 50)"}},
            },
        ),
        
        # ===== ENTRA ID GOVERNANCE - ACCESS REVIEWS =====
        Tool(
            name="list_access_reviews",
            description="List all access review definitions",
            inputSchema={"type": "object", "properties": {}},
        ),
        Tool(
            name="get_access_review",
            description="Get details of a specific access review",
            inputSchema={
                "type": "object",
                "properties": {"review_id": {"type": "string", "description": "Access review definition ID"}},
                "required": ["review_id"],
            },
        ),
        Tool(
            name="create_access_review",
            description="Create a new access review for a group or application",
            inputSchema={
                "type": "object",
                "properties": {
                    "display_name": {"type": "string", "description": "Display name for the access review"},
                    "description": {"type": "string", "description": "Description of the access review"},
                    "group_id": {"type": "string", "description": "Group ID to review (optional)"},
                    "reviewers": {"type": "array", "description": "Array of reviewer user IDs", "items": {"type": "string"}},
                    "duration_days": {"type": "number", "description": "Duration of the review in days (default: 14)"}
                },
                "required": ["display_name"],
            },
        ),
        Tool(
            name="start_access_review",
            description="Start an access review instance",
            inputSchema={
                "type": "object",
                "properties": {"review_id": {"type": "string", "description": "Access review definition ID"}},
                "required": ["review_id"],
            },
        ),
        Tool(
            name="stop_access_review",
            description="Stop an access review instance",
            inputSchema={
                "type": "object",
                "properties": {
                    "review_id": {"type": "string", "description": "Access review definition ID"},
                    "instance_id": {"type": "string", "description": "Access review instance ID"}
                },
                "required": ["review_id", "instance_id"],
            },
        ),
        Tool(
            name="get_access_review_decisions",
            description="Get decisions from an access review",
            inputSchema={
                "type": "object",
                "properties": {
                    "review_id": {"type": "string", "description": "Access review definition ID"},
                    "instance_id": {"type": "string", "description": "Access review instance ID"}
                },
                "required": ["review_id", "instance_id"],
            },
        ),
        
        # ===== ENTRA ID GOVERNANCE - ENTITLEMENT MANAGEMENT =====
        Tool(
            name="list_access_packages",
            description="List all access packages in entitlement management",
            inputSchema={"type": "object", "properties": {}},
        ),
        Tool(
            name="get_access_package",
            description="Get details of a specific access package",
            inputSchema={
                "type": "object",
                "properties": {"package_id": {"type": "string", "description": "Access package ID"}},
                "required": ["package_id"],
            },
        ),
        Tool(
            name="create_access_package",
            description="Create a new access package",
            inputSchema={
                "type": "object",
                "properties": {
                    "display_name": {"type": "string", "description": "Display name for the access package"},
                    "description": {"type": "string", "description": "Description of the access package"},
                    "catalog_id": {"type": "string", "description": "Catalog ID where the package will be created"}
                },
                "required": ["display_name", "catalog_id"],
            },
        ),
        Tool(
            name="list_access_package_assignments",
            description="List all access package assignments",
            inputSchema={
                "type": "object",
                "properties": {"package_id": {"type": "string", "description": "Access package ID (optional)"}},
            },
        ),
        Tool(
            name="request_access_package",
            description="Create an access package assignment request",
            inputSchema={
                "type": "object",
                "properties": {
                    "package_id": {"type": "string", "description": "Access package ID"},
                    "user_id": {"type": "string", "description": "User ID requesting access"},
                    "justification": {"type": "string", "description": "Justification for the request"}
                },
                "required": ["package_id", "user_id"],
            },
        ),
        Tool(
            name="list_catalogs",
            description="List all entitlement management catalogs",
            inputSchema={"type": "object", "properties": {}},
        ),
        Tool(
            name="create_catalog",
            description="Create a new entitlement management catalog",
            inputSchema={
                "type": "object",
                "properties": {
                    "display_name": {"type": "string", "description": "Display name for the catalog"},
                    "description": {"type": "string", "description": "Description of the catalog"}
                },
                "required": ["display_name"],
            },
        ),
        
        # ===== ENTRA ID GOVERNANCE - PRIVILEGED IDENTITY MANAGEMENT (PIM) =====
        Tool(
            name="list_privileged_roles",
            description="List all privileged roles available for PIM",
            inputSchema={"type": "object", "properties": {}},
        ),
        Tool(
            name="get_privileged_role_assignments",
            description="Get privileged role assignments for a user",
            inputSchema={
                "type": "object",
                "properties": {"user_id": {"type": "string", "description": "User ID (object ID) or User Principal Name"}},
                "required": ["user_id"],
            },
        ),
        Tool(
            name="list_eligible_role_assignments",
            description="List all eligible role assignments (PIM)",
            inputSchema={"type": "object", "properties": {}},
        ),
        Tool(
            name="create_eligible_role_assignment",
            description="Create an eligible role assignment in PIM",
            inputSchema={
                "type": "object",
                "properties": {
                    "user_id": {"type": "string", "description": "User ID to assign the role to"},
                    "role_definition_id": {"type": "string", "description": "Role definition ID"},
                    "justification": {"type": "string", "description": "Justification for the assignment"},
                    "duration_hours": {"type": "number", "description": "Duration in hours (max: 8760 = 1 year)"}
                },
                "required": ["user_id", "role_definition_id"],
            },
        ),
        Tool(
            name="activate_role",
            description="Activate an eligible role assignment",
            inputSchema={
                "type": "object",
                "properties": {
                    "user_id": {"type": "string", "description": "User ID"},
                    "role_definition_id": {"type": "string", "description": "Role definition ID"},
                    "justification": {"type": "string", "description": "Justification for activation"},
                    "duration_hours": {"type": "number", "description": "Activation duration in hours (default: 8)"}
                },
                "required": ["user_id", "role_definition_id", "justification"],
            },
        ),
        Tool(
            name="deactivate_role",
            description="Deactivate an active role assignment",
            inputSchema={
                "type": "object",
                "properties": {
                    "assignment_id": {"type": "string", "description": "Role assignment schedule request ID"}
                },
                "required": ["assignment_id"],
            },
        ),
        
        # ===== ENTRA ID GOVERNANCE - LIFECYCLE WORKFLOWS =====
        Tool(
            name="list_lifecycle_workflows",
            description="List all lifecycle workflows",
            inputSchema={"type": "object", "properties": {}},
        ),
        Tool(
            name="get_lifecycle_workflow",
            description="Get details of a specific lifecycle workflow",
            inputSchema={
                "type": "object",
                "properties": {"workflow_id": {"type": "string", "description": "Workflow ID"}},
                "required": ["workflow_id"],
            },
        ),
        Tool(
            name="create_lifecycle_workflow",
            description="Create a new lifecycle workflow",
            inputSchema={
                "type": "object",
                "properties": {
                    "display_name": {"type": "string", "description": "Display name for the workflow"},
                    "description": {"type": "string", "description": "Description of the workflow"},
                    "category": {"type": "string", "description": "Category: 'joiner', 'leaver', or 'mover'", "enum": ["joiner", "leaver", "mover"]},
                    "trigger_type": {"type": "string", "description": "When to trigger: 'onDemand' or 'scheduled'"}
                },
                "required": ["display_name", "category"],
            },
        ),
        Tool(
            name="enable_lifecycle_workflow",
            description="Enable a lifecycle workflow",
            inputSchema={
                "type": "object",
                "properties": {"workflow_id": {"type": "string", "description": "Workflow ID"}},
                "required": ["workflow_id"],
            },
        ),
        Tool(
            name="disable_lifecycle_workflow",
            description="Disable a lifecycle workflow",
            inputSchema={
                "type": "object",
                "properties": {"workflow_id": {"type": "string", "description": "Workflow ID"}},
                "required": ["workflow_id"],
            },
        ),
        Tool(
            name="run_lifecycle_workflow",
            description="Run a lifecycle workflow on-demand for specific users",
            inputSchema={
                "type": "object",
                "properties": {
                    "workflow_id": {"type": "string", "description": "Workflow ID"},
                    "user_ids": {"type": "array", "description": "Array of user IDs to run workflow for", "items": {"type": "string"}}
                },
                "required": ["workflow_id", "user_ids"],
            },
        ),
        Tool(
            name="get_workflow_execution_history",
            description="Get execution history for a lifecycle workflow",
            inputSchema={
                "type": "object",
                "properties": {
                    "workflow_id": {"type": "string", "description": "Workflow ID"},
                    "top": {"type": "number", "description": "Number of results (default: 50)"}
                },
                "required": ["workflow_id"],
            },
        ),
    ]


@server.call_tool()
async def handle_call_tool(name: str, arguments: dict | None) -> list[TextContent]:
    """Handle tool execution requests"""
    
    if arguments is None:
        arguments = {}
    
    try:
        # Initialize Graph client
        client = await initialize_graph_client()
        
        # ===== ENTRA ID OPERATIONS =====
        if name == "list_users":
            top = arguments.get("top", 10)
            filter_query = arguments.get("filter")
            search_query = arguments.get("search")
            
            request_config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
                query_parameters=UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
                    top=min(int(top), 999),
                    filter=filter_query,
                    search=search_query,
                    select=["id", "displayName", "userPrincipalName", "mail", "jobTitle", "department"]
                )
            )
            
            users = await client.users.get(request_configuration=request_config)
            
            user_list = []
            if users and users.value:
                for user in users.value:
                    user_list.append({
                        "id": user.id,
                        "displayName": user.display_name,
                        "userPrincipalName": user.user_principal_name,
                        "mail": user.mail,
                        "jobTitle": user.job_title,
                        "department": user.department
                    })
            
            return [TextContent(type="text", text=json.dumps(user_list, indent=2))]
        
        elif name == "get_user":
            user_id = arguments["user_id"]
            user = await client.users.by_user_id(user_id).get()
            
            user_data = {
                "id": user.id,
                "displayName": user.display_name,
                "userPrincipalName": user.user_principal_name,
                "mail": user.mail,
                "jobTitle": user.job_title,
                "department": user.department,
                "officeLocation": user.office_location,
                "mobilePhone": user.mobile_phone,
                "businessPhones": user.business_phones,
                "accountEnabled": user.account_enabled
            }
            
            return [TextContent(type="text", text=json.dumps(user_data, indent=2))]
        
        elif name == "list_groups":
            top = arguments.get("top", 10)
            filter_query = arguments.get("filter")
            
            request_config = GroupsRequestBuilder.GroupsRequestBuilderGetRequestConfiguration(
                query_parameters=GroupsRequestBuilder.GroupsRequestBuilderGetQueryParameters(
                    top=int(top),
                    filter=filter_query,
                    select=["id", "displayName", "description", "mail", "groupTypes"]
                )
            )
            
            groups = await client.groups.get(request_configuration=request_config)
            
            group_list = []
            if groups and groups.value:
                for group in groups.value:
                    group_list.append({
                        "id": group.id,
                        "displayName": group.display_name,
                        "description": group.description,
                        "mail": group.mail,
                        "groupTypes": group.group_types
                    })
            
            return [TextContent(type="text", text=json.dumps(group_list, indent=2))]
        
        elif name == "get_group":
            group_id = arguments["group_id"]
            group = await client.groups.by_group_id(group_id).get()
            
            group_data = {
                "id": group.id,
                "displayName": group.display_name,
                "description": group.description,
                "mail": group.mail,
                "groupTypes": group.group_types,
                "mailEnabled": group.mail_enabled,
                "securityEnabled": group.security_enabled
            }
            
            return [TextContent(type="text", text=json.dumps(group_data, indent=2))]
        
        elif name == "get_group_members":
            group_id = arguments["group_id"]
            top = arguments.get("top", 10)
            
            members = await client.groups.by_group_id(group_id).members.get()
            
            member_list = []
            if members and members.value:
                for member in members.value[:int(top)]:
                    member_list.append({
                        "id": member.id,
                        "displayName": getattr(member, 'display_name', 'N/A'),
                        "userPrincipalName": getattr(member, 'user_principal_name', 'N/A')
                    })
            
            return [TextContent(type="text", text=json.dumps(member_list, indent=2))]
        
        elif name == "get_user_groups":
            user_id = arguments["user_id"]
            groups = await client.users.by_user_id(user_id).member_of.get()
            
            group_list = []
            if groups and groups.value:
                for group in groups.value:
                    group_list.append({
                        "id": group.id,
                        "displayName": getattr(group, 'display_name', 'N/A'),
                        "groupTypes": getattr(group, 'group_types', [])
                    })
            
            return [TextContent(type="text", text=json.dumps(group_list, indent=2))]
        
        elif name == "search_users":
            query = arguments["query"]
            top = arguments.get("top", 10)
            
            filter_query = f"startswith(displayName,'{query}') or startswith(mail,'{query}') or startswith(userPrincipalName,'{query}')"
            
            request_config = UsersRequestBuilder.UsersRequestBuilderGetRequestConfiguration(
                query_parameters=UsersRequestBuilder.UsersRequestBuilderGetQueryParameters(
                    top=int(top),
                    filter=filter_query,
                    select=["id", "displayName", "userPrincipalName", "mail", "jobTitle"]
                )
            )
            
            users = await client.users.get(request_configuration=request_config)
            
            user_list = []
            if users and users.value:
                for user in users.value:
                    user_list.append({
                        "id": user.id,
                        "displayName": user.display_name,
                        "userPrincipalName": user.user_principal_name,
                        "mail": user.mail,
                        "jobTitle": user.job_title
                    })
            
            return [TextContent(type="text", text=json.dumps(user_list, indent=2))]
        
        elif name == "get_user_manager":
            user_id = arguments["user_id"]
            manager = await client.users.by_user_id(user_id).manager.get()
            
            manager_data = {
                "id": manager.id,
                "displayName": getattr(manager, 'display_name', 'N/A'),
                "userPrincipalName": getattr(manager, 'user_principal_name', 'N/A'),
                "mail": getattr(manager, 'mail', 'N/A'),
                "jobTitle": getattr(manager, 'job_title', 'N/A')
            }
            
            return [TextContent(type="text", text=json.dumps(manager_data, indent=2))]
        
        elif name == "get_user_direct_reports":
            user_id = arguments["user_id"]
            reports = await client.users.by_user_id(user_id).direct_reports.get()
            
            report_list = []
            if reports and reports.value:
                for report in reports.value:
                    report_list.append({
                        "id": report.id,
                        "displayName": getattr(report, 'display_name', 'N/A'),
                        "userPrincipalName": getattr(report, 'user_principal_name', 'N/A'),
                        "jobTitle": getattr(report, 'job_title', 'N/A')
                    })
            
            return [TextContent(type="text", text=json.dumps(report_list, indent=2))]
        
        # ===== INTERNET ACCESS OPERATIONS =====
        elif name == "get_internet_access_status":
            result = await make_graph_request('GET', 'networkaccess/settings')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "list_forwarding_profiles":
            result = await make_graph_request('GET', 'networkaccess/forwardingProfiles')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "get_forwarding_profile":
            profile_id = arguments["profile_id"]
            result = await make_graph_request('GET', f'networkaccess/forwardingProfiles/{profile_id}')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "create_forwarding_profile":
            profile_data = {
                "name": arguments["name"],
                "description": arguments.get("description", ""),
                "trafficForwardingType": arguments["profile_type"],
                "priority": arguments.get("priority", 100),
                "state": "enabled"
            }
            
            result = await make_graph_request('POST', 'networkaccess/forwardingProfiles', profile_data)
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "update_forwarding_profile":
            profile_id = arguments["profile_id"]
            
            update_data = {}
            if "name" in arguments:
                update_data["name"] = arguments["name"]
            if "description" in arguments:
                update_data["description"] = arguments["description"]
            if "state" in arguments:
                update_data["state"] = arguments["state"]
            
            result = await make_graph_request('PATCH', f'networkaccess/forwardingProfiles/{profile_id}', update_data)
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "delete_forwarding_profile":
            profile_id = arguments["profile_id"]
            result = await make_graph_request('DELETE', f'networkaccess/forwardingProfiles/{profile_id}')
            return [TextContent(type="text", text="Forwarding profile deleted successfully")]
        
        # ===== SECURITY POLICIES =====
        elif name == "list_security_policies":
            result = await make_graph_request('GET', 'networkaccess/filteringPolicies')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "get_security_policy":
            policy_id = arguments["policy_id"]
            result = await make_graph_request('GET', f'networkaccess/filteringPolicies/{policy_id}')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "create_security_policy":
            policy_data = {
                "name": arguments["name"],
                "description": arguments.get("description", ""),
                "priority": arguments.get("priority", 100),
                "action": arguments["action"],
                "state": "enabled"
            }
            
            result = await make_graph_request('POST', 'networkaccess/filteringPolicies', policy_data)
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "add_fqdn_to_policy":
            policy_id = arguments["policy_id"]
            fqdn = arguments["fqdn"]
            
            result = await make_graph_request('POST', f'networkaccess/filteringPolicies/{policy_id}/policyRules', {
                "@odata.type": "#microsoft.graph.networkaccess.fqdnFilteringRule",
                "fqdn": fqdn
            })
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "add_web_category_to_policy":
            policy_id = arguments["policy_id"]
            category = arguments["category"]
            
            result = await make_graph_request('POST', f'networkaccess/filteringPolicies/{policy_id}/policyRules', {
                "@odata.type": "#microsoft.graph.networkaccess.webCategoryFilteringRule",
                "webCategory": category
            })
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "update_security_policy":
            policy_id = arguments["policy_id"]
            
            update_data = {}
            if "name" in arguments:
                update_data["name"] = arguments["name"]
            if "description" in arguments:
                update_data["description"] = arguments["description"]
            if "state" in arguments:
                update_data["state"] = arguments["state"]
            
            result = await make_graph_request('PATCH', f'networkaccess/filteringPolicies/{policy_id}', update_data)
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "delete_security_policy":
            policy_id = arguments["policy_id"]
            result = await make_graph_request('DELETE', f'networkaccess/filteringPolicies/{policy_id}')
            return [TextContent(type="text", text="Security policy deleted successfully")]
        
        # ===== PRIVATE ACCESS OPERATIONS =====
        elif name == "list_private_access_apps":
            result = await make_graph_request('GET', 'networkaccess/connectivity/branches')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "get_private_access_app":
            app_id = arguments["app_id"]
            result = await make_graph_request('GET', f'networkaccess/connectivity/branches/{app_id}')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "create_private_access_app":
            app_data = {
                "name": arguments["name"],
                "description": arguments.get("description", ""),
                "region": "usEast",
                "deviceLinks": []
            }
            
            segments = []
            
            if "fqdn" in arguments:
                segments.append({
                    "@odata.type": "#microsoft.graph.networkaccess.fqdnSegment",
                    "fqdn": arguments["fqdn"]
                })
            
            if "ip_addresses" in arguments:
                for ip in arguments["ip_addresses"]:
                    segments.append({
                        "@odata.type": "#microsoft.graph.networkaccess.ipAddressSegment",
                        "ipAddress": ip
                    })
            
            if segments:
                app_data["deviceLinks"] = [{
                    "name": f"{arguments['name']}-link",
                    "deviceVendor": "other",
                    "bandwidthCapacityInMbps": "mbps1000",
                    "bgpConfiguration": {
                        "asn": 65000,
                        "ipAddress": "10.0.0.1"
                    },
                    "tunnelConfiguration": {
                        "@odata.type": "#microsoft.graph.networkaccess.tunnelConfigurationIKEv2Default"
                    }
                }]
            
            result = await make_graph_request('POST', 'networkaccess/connectivity/branches', app_data)
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "assign_users_to_private_access_app":
            app_id = arguments["app_id"]
            
            assignments = []
            
            if "user_ids" in arguments:
                for user_id in arguments["user_ids"]:
                    assignments.append({
                        "@odata.type": "#microsoft.graph.networkaccess.userAssignment",
                        "userId": user_id
                    })
            
            if "group_ids" in arguments:
                for group_id in arguments["group_ids"]:
                    assignments.append({
                        "@odata.type": "#microsoft.graph.networkaccess.groupAssignment",
                        "groupId": group_id
                    })
            
            result = {"message": "User/group assignments configured", "assignments": assignments}
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "update_private_access_app":
            app_id = arguments["app_id"]
            
            update_data = {}
            if "name" in arguments:
                update_data["name"] = arguments["name"]
            if "description" in arguments:
                update_data["description"] = arguments["description"]
            
            result = await make_graph_request('PATCH', f'networkaccess/connectivity/branches/{app_id}', update_data)
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "delete_private_access_app":
            app_id = arguments["app_id"]
            result = await make_graph_request('DELETE', f'networkaccess/connectivity/branches/{app_id}')
            return [TextContent(type="text", text="Private Access app deleted successfully")]
        
        elif name == "list_private_access_connectors":
            result = await make_graph_request('GET', 'networkaccess/connectivity/remoteNetworks')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "get_private_access_connector":
            connector_id = arguments["connector_id"]
            result = await make_graph_request('GET', f'networkaccess/connectivity/remoteNetworks/{connector_id}')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        # ===== TRAFFIC LOGS =====
        elif name == "get_traffic_logs":
            top = arguments.get("top", 50)
            result = await make_graph_request('GET', f'networkaccess/logs/traffic?$top={top}')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        # ===== ACCESS REVIEWS =====
        elif name == "list_access_reviews":
            result = await make_graph_request('GET', 'identityGovernance/accessReviews/definitions')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "get_access_review":
            review_id = arguments["review_id"]
            result = await make_graph_request('GET', f'identityGovernance/accessReviews/definitions/{review_id}')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "create_access_review":
            duration = arguments.get("duration_days", 14)
            end_date = (datetime.now() + timedelta(days=duration)).isoformat() + "Z"
            
            review_data = {
                "displayName": arguments["display_name"],
                "descriptionForAdmins": arguments.get("description", ""),
                "descriptionForReviewers": arguments.get("description", ""),
                "scope": {
                    "@odata.type": "#microsoft.graph.accessReviewQueryScope",
                    "query": f"/groups/{arguments.get('group_id')}/members" if "group_id" in arguments else "/users",
                    "queryType": "MicrosoftGraph"
                },
                "reviewers": [],
                "settings": {
                    "mailNotificationsEnabled": True,
                    "reminderNotificationsEnabled": True,
                    "justificationRequiredOnApproval": True,
                    "defaultDecisionEnabled": False,
                    "defaultDecision": "None",
                    "instanceDurationInDays": duration,
                    "recurrence": {
                        "pattern": {
                            "type": "absoluteMonthly",
                            "interval": 3
                        },
                        "range": {
                            "type": "noEnd",
                            "startDate": datetime.now().strftime("%Y-%m-%d")
                        }
                    }
                }
            }
            
            if "reviewers" in arguments:
                for reviewer_id in arguments["reviewers"]:
                    review_data["reviewers"].append({
                        "query": f"/users/{reviewer_id}",
                        "queryType": "MicrosoftGraph"
                    })
            
            result = await make_graph_request('POST', 'identityGovernance/accessReviews/definitions', review_data)
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "start_access_review":
            review_id = arguments["review_id"]
            result = await make_graph_request('POST', f'identityGovernance/accessReviews/definitions/{review_id}/instances', {})
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "stop_access_review":
            review_id = arguments["review_id"]
            instance_id = arguments["instance_id"]
            result = await make_graph_request('POST', f'identityGovernance/accessReviews/definitions/{review_id}/instances/{instance_id}/stop', {})
            return [TextContent(type="text", text="Access review stopped successfully")]
        
        elif name == "get_access_review_decisions":
            review_id = arguments["review_id"]
            instance_id = arguments["instance_id"]
            result = await make_graph_request('GET', f'identityGovernance/accessReviews/definitions/{review_id}/instances/{instance_id}/decisions')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        # ===== ENTITLEMENT MANAGEMENT =====
        elif name == "list_access_packages":
            result = await make_graph_request('GET', 'identityGovernance/entitlementManagement/accessPackages')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "get_access_package":
            package_id = arguments["package_id"]
            result = await make_graph_request('GET', f'identityGovernance/entitlementManagement/accessPackages/{package_id}')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "create_access_package":
            package_data = {
                "displayName": arguments["display_name"],
                "description": arguments.get("description", ""),
                "catalogId": arguments["catalog_id"],
                "isHidden": False
            }
            
            result = await make_graph_request('POST', 'identityGovernance/entitlementManagement/accessPackages', package_data)
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "list_access_package_assignments":
            if "package_id" in arguments:
                result = await make_graph_request('GET', f'identityGovernance/entitlementManagement/accessPackageAssignments?$filter=accessPackageId eq \'{arguments["package_id"]}\'')
            else:
                result = await make_graph_request('GET', 'identityGovernance/entitlementManagement/accessPackageAssignments')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "request_access_package":
            request_data = {
                "requestType": "UserAdd",
                "accessPackageAssignment": {
                    "targetId": arguments["user_id"],
                    "assignmentPolicyId": arguments["package_id"],
                    "accessPackageId": arguments["package_id"]
                },
                "answers": []
            }
            
            if "justification" in arguments:
                request_data["justification"] = arguments["justification"]
            
            result = await make_graph_request('POST', 'identityGovernance/entitlementManagement/accessPackageAssignmentRequests', request_data)
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "list_catalogs":
            result = await make_graph_request('GET', 'identityGovernance/entitlementManagement/catalogs')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "create_catalog":
            catalog_data = {
                "displayName": arguments["display_name"],
                "description": arguments.get("description", ""),
                "isExternallyVisible": False
            }
            
            result = await make_graph_request('POST', 'identityGovernance/entitlementManagement/catalogs', catalog_data)
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        # ===== PRIVILEGED IDENTITY MANAGEMENT =====
        elif name == "list_privileged_roles":
            result = await make_graph_request('GET', 'roleManagement/directory/roleDefinitions')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "get_privileged_role_assignments":
            user_id = arguments["user_id"]
            result = await make_graph_request('GET', f'roleManagement/directory/roleAssignments?$filter=principalId eq \'{user_id}\'')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "list_eligible_role_assignments":
            result = await make_graph_request('GET', 'roleManagement/directory/roleEligibilitySchedules')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "create_eligible_role_assignment":
            duration = arguments.get("duration_hours", 8760)  # Default 1 year
            
            assignment_data = {
                "action": "adminAssign",
                "justification": arguments.get("justification", "PIM eligible role assignment"),
                "roleDefinitionId": arguments["role_definition_id"],
                "directoryScopeId": "/",
                "principalId": arguments["user_id"],
                "scheduleInfo": {
                    "startDateTime": datetime.now().isoformat() + "Z",
                    "expiration": {
                        "type": "afterDuration",
                        "duration": f"PT{duration}H"
                    }
                }
            }
            
            result = await make_graph_request('POST', 'roleManagement/directory/roleEligibilityScheduleRequests', assignment_data)
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "activate_role":
            duration = arguments.get("duration_hours", 8)
            
            activation_data = {
                "action": "selfActivate",
                "justification": arguments["justification"],
                "roleDefinitionId": arguments["role_definition_id"],
                "directoryScopeId": "/",
                "principalId": arguments["user_id"],
                "scheduleInfo": {
                    "startDateTime": datetime.now().isoformat() + "Z",
                    "expiration": {
                        "type": "afterDuration",
                        "duration": f"PT{duration}H"
                    }
                }
            }
            
            result = await make_graph_request('POST', 'roleManagement/directory/roleAssignmentScheduleRequests', activation_data)
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "deactivate_role":
            assignment_id = arguments["assignment_id"]
            result = await make_graph_request('DELETE', f'roleManagement/directory/roleAssignmentScheduleRequests/{assignment_id}')
            return [TextContent(type="text", text="Role deactivated successfully")]
        
        # ===== LIFECYCLE WORKFLOWS =====
        elif name == "list_lifecycle_workflows":
            result = await make_graph_request('GET', 'identityGovernance/lifecycleWorkflows/workflows')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "get_lifecycle_workflow":
            workflow_id = arguments["workflow_id"]
            result = await make_graph_request('GET', f'identityGovernance/lifecycleWorkflows/workflows/{workflow_id}')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "create_lifecycle_workflow":
            workflow_data = {
                "displayName": arguments["display_name"],
                "description": arguments.get("description", ""),
                "category": arguments["category"],
                "isEnabled": False,
                "isSchedulingEnabled": arguments.get("trigger_type", "onDemand") == "scheduled",
                "executionConditions": {
                    "@odata.type": "#microsoft.graph.identityGovernance.triggerAndScopeBasedConditions",
                    "scope": {
                        "@odata.type": "#microsoft.graph.identityGovernance.ruleBasedSubjectSet",
                        "rule": "department eq 'Sales'"
                    },
                    "trigger": {
                        "@odata.type": "#microsoft.graph.identityGovernance.timeBasedAttributeTrigger",
                        "timeBasedAttribute": "employeeHireDate",
                        "offsetInDays": 0
                    }
                },
                "tasks": []
            }
            
            result = await make_graph_request('POST', 'identityGovernance/lifecycleWorkflows/workflows', workflow_data)
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "enable_lifecycle_workflow":
            workflow_id = arguments["workflow_id"]
            result = await make_graph_request('PATCH', f'identityGovernance/lifecycleWorkflows/workflows/{workflow_id}', {"isEnabled": True})
            return [TextContent(type="text", text="Workflow enabled successfully")]
        
        elif name == "disable_lifecycle_workflow":
            workflow_id = arguments["workflow_id"]
            result = await make_graph_request('PATCH', f'identityGovernance/lifecycleWorkflows/workflows/{workflow_id}', {"isEnabled": False})
            return [TextContent(type="text", text="Workflow disabled successfully")]
        
        elif name == "run_lifecycle_workflow":
            workflow_id = arguments["workflow_id"]
            user_ids = arguments["user_ids"]
            
            run_data = {
                "subjects": [{"id": user_id} for user_id in user_ids]
            }
            
            result = await make_graph_request('POST', f'identityGovernance/lifecycleWorkflows/workflows/{workflow_id}/activate', run_data)
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        elif name == "get_workflow_execution_history":
            workflow_id = arguments["workflow_id"]
            top = arguments.get("top", 50)
            result = await make_graph_request('GET', f'identityGovernance/lifecycleWorkflows/workflows/{workflow_id}/runs?$top={top}')
            return [TextContent(type="text", text=json.dumps(result, indent=2))]
        
        else:
            raise ValueError(f"Unknown tool: {name}")
    
    except Exception as e:
        return [TextContent(type="text", text=f"Error: {str(e)}")]


async def main():
    """Run the MCP server"""
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            InitializationOptions(
                server_name="mcp-entra-suite",
                server_version="1.0.0",
                capabilities=server.get_capabilities(
                    notification_options=NotificationOptions(),
                    experimental_capabilities={},
                ),
            ),
        )


if __name__ == "__main__":
    asyncio.run(main())