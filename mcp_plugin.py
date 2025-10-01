"""
Excel MCP Plugin - Main Module
Provides Excel integration via xlwings
"""

import xlwings as xw
from mcp_connector import MCPConnector
from sheet_manager import SheetManager
import json
from typing import Optional


# Initialize connector (global instance)
connector = MCPConnector()


@xw.func
def mcp_call_tool(server: str, tool: str, arguments: str = "{}") -> str:
    """
    Call an MCP tool and return the result
    
    Args:
        server: MCP server name
        tool: Tool name
        arguments: JSON string of arguments (default: "{}")
        
    Returns:
        Tool execution result as string
        
    Example:
        =MCP_CALL_TOOL("weather-mcp", "get_weather", "{""city"": ""London""}")
    """
    try:
        # Parse arguments
        args_dict = json.loads(arguments)
        
        # Call tool
        result = connector.call_tool(server, tool, args_dict)
        
        if result["success"]:
            # Return formatted result
            if isinstance(result["data"], dict):
                return json.dumps(result["data"], indent=2)
            else:
                return str(result.get("raw_output", result["data"]))
        else:
            return f"ERROR: {result.get('error', 'Unknown error')}"
            
    except json.JSONDecodeError:
        return "ERROR: Invalid JSON arguments"
    except Exception as e:
        return f"ERROR: {str(e)}"


@xw.func
def mcp_read_resource(server: str, resource_uri: str) -> str:
    """
    Read an MCP resource and return its content
    
    Args:
        server: MCP server name
        resource_uri: Resource URI
        
    Returns:
        Resource content as string
        
    Example:
        =MCP_READ_RESOURCE("database-mcp", "db://schema")
    """
    try:
        result = connector.read_resource(server, resource_uri)
        
        if result["success"]:
            if isinstance(result["data"], dict):
                return json.dumps(result["data"], indent=2)
            else:
                return str(result.get("raw_output", result["data"]))
        else:
            return f"ERROR: {result.get('error', 'Unknown error')}"
            
    except Exception as e:
        return f"ERROR: {str(e)}"


@xw.func
def mcp_get_prompt(server: str, prompt: str, arguments: str = "{}") -> str:
    """
    Get an MCP prompt with arguments filled in
    
    Args:
        server: MCP server name
        prompt: Prompt name
        arguments: JSON string of arguments (default: "{}")
        
    Returns:
        Prompt template with arguments filled in
        
    Example:
        =MCP_GET_PROMPT("travel-mcp", "plan_itinerary", "{""city"": ""Paris""}")
    """
    try:
        # Parse arguments
        args_dict = json.loads(arguments)
        
        # Call prompt
        result = connector.call_prompt(server, prompt, args_dict)
        
        if result["success"]:
            if isinstance(result["data"], dict):
                return json.dumps(result["data"], indent=2)
            else:
                return str(result.get("raw_output", result["data"]))
        else:
            return f"ERROR: {result.get('error', 'Unknown error')}"
            
    except json.JSONDecodeError:
        return "ERROR: Invalid JSON arguments"
    except Exception as e:
        return f"ERROR: {str(e)}"


@xw.func
def mcp_server_status(server: str) -> str:
    """
    Check the connection status of an MCP server
    
    Args:
        server: MCP server name
        
    Returns:
        Connection status ("Connected", "Disconnected", or "Error")
        
    Example:
        =MCP_SERVER_STATUS("weather-mcp")
    """
    try:
        server_obj = connector.check_server_connection(server)
        return server_obj.status
    except Exception as e:
        return f"ERROR: {str(e)}"


@xw.func
def mcp_list_servers() -> str:
    """
    List all available MCP servers
    
    Returns:
        Comma-separated list of server names
        
    Example:
        =MCP_LIST_SERVERS()
    """
    try:
        servers = connector.discover_servers()
        return ", ".join([s.name for s in servers])
    except Exception as e:
        return f"ERROR: {str(e)}"


def refresh_mcps_sheet():
    """
    Refresh the MCPs sheet with current server data
    Called from Excel button or VBA
    """
    try:
        wb = xw.Book.caller()
        manager = SheetManager(wb, connector)
        
        # Update sheet
        servers_found, servers_connected, errors = manager.update_sheet()
        
        # Show result message
        if errors:
            error_msg = "\n".join(errors)
            message = f"Found {servers_found} servers, {servers_connected} connected.\n\nErrors:\n{error_msg}"
        else:
            message = f"Successfully updated!\nFound {servers_found} servers, {servers_connected} connected."
        
        # Show message box
        wb.app.api.MessageBox(message, "MCP Sheet Update")
        
    except Exception as e:
        wb = xw.Book.caller()
        wb.app.api.MessageBox(f"Error updating MCPs sheet:\n{str(e)}", "Error")


def initialize_mcps_sheet():
    """
    Initialize a new MCPs sheet
    Called from Excel button or VBA
    """
    try:
        wb = xw.Book.caller()
        manager = SheetManager(wb, connector)
        
        # Ensure sheet exists
        sheet = manager.ensure_sheet_exists()
        
        wb.app.api.MessageBox(
            f"MCPs sheet initialized successfully!\n\nUse 'Refresh MCPs' to populate with server data.",
            "MCP Sheet Initialized"
        )
        
    except Exception as e:
        wb = xw.Book.caller()
        wb.app.api.MessageBox(f"Error initializing MCPs sheet:\n{str(e)}", "Error")


def clear_mcps_sheet():
    """
    Clear the MCPs sheet
    Called from Excel button or VBA
    """
    try:
        wb = xw.Book.caller()
        manager = SheetManager(wb, connector)
        
        # Confirm with user
        response = wb.app.api.MessageBox(
            "Are you sure you want to clear the MCPs sheet?",
            "Confirm Clear",
            4  # Yes/No buttons
        )
        
        if response == 6:  # Yes button
            manager.clear_sheet()
            wb.app.api.MessageBox("MCPs sheet cleared successfully!", "Sheet Cleared")
        
    except Exception as e:
        wb = xw.Book.caller()
        wb.app.api.MessageBox(f"Error clearing MCPs sheet:\n{str(e)}", "Error")


def show_capability_details():
    """
    Show detailed information about the selected capability
    Called from Excel button or VBA
    """
    try:
        wb = xw.Book.caller()
        manager = SheetManager(wb, connector)
        
        # Get active cell
        active_cell = wb.app.selection
        row = active_cell.row
        col = active_cell.column
        
        # Get sheet
        try:
            sheet = wb.sheets[SheetManager.SHEET_NAME]
        except KeyError:
            wb.app.api.MessageBox("MCPs sheet not found!", "Error")
            return
        
        # Get capability info
        capability_info = manager.get_capability_at_cell(sheet, row, col)
        
        if not capability_info:
            wb.app.api.MessageBox(
                "Please select a capability cell in the MCPs sheet.",
                "No Capability Selected"
            )
            return
        
        server_name, capability_type, capability_name = capability_info
        
        # Get details based on type
        if capability_type == "tool":
            details = connector.get_tool_details(server_name, capability_name)
        elif capability_type == "resource":
            details = connector.get_resource_details(server_name, capability_name)
        elif capability_type == "prompt":
            details = connector.get_prompt_details(server_name, capability_name)
        else:
            details = None
        
        if details:
            details_str = json.dumps(details, indent=2)
            message = f"Server: {server_name}\nType: {capability_type}\nName: {capability_name}\n\nDetails:\n{details_str}"
        else:
            message = f"Server: {server_name}\nType: {capability_type}\nName: {capability_name}\n\nNo additional details available."
        
        # Show details (truncate if too long)
        if len(message) > 500:
            message = message[:500] + "\n\n... (truncated)"
        
        wb.app.api.MessageBox(message, "Capability Details")
        
    except Exception as e:
        wb = xw.Book.caller()
        wb.app.api.MessageBox(f"Error getting capability details:\n{str(e)}", "Error")


def main():
    """
    Main function - called by the 'Run main' button
    Refreshes the MCPs sheet
    """
    refresh_mcps_sheet()


if __name__ == "__main__":
    # For testing purposes
    print("Excel MCP Plugin loaded successfully!")
    print("Available functions:")
    print("  - refresh_mcps_sheet()")
    print("  - initialize_mcps_sheet()")
    print("  - clear_mcps_sheet()")
    print("  - show_capability_details()")
    print("\nAvailable UDFs:")
    print("  - MCP_CALL_TOOL(server, tool, arguments)")
    print("  - MCP_READ_RESOURCE(server, resource_uri)")
    print("  - MCP_GET_PROMPT(server, prompt, arguments)")
    print("  - MCP_SERVER_STATUS(server)")
    print("  - MCP_LIST_SERVERS()")
