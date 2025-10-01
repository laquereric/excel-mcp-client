"""
MCP Connector Module
Handles communication with MCP servers via manus-mcp-cli
"""

import subprocess
import json
from typing import List, Dict, Optional, Any
from dataclasses import dataclass


@dataclass
class MCPServer:
    """Represents an MCP server"""
    name: str
    status: str  # "Connected", "Disconnected", "Error"
    error_message: Optional[str] = None


@dataclass
class MCPCapability:
    """Represents a capability (tool, resource, or prompt)"""
    name: str
    capability_type: str  # "tool", "resource", "prompt"
    description: Optional[str] = None
    details: Optional[Dict[str, Any]] = None


class MCPConnector:
    """Manages connections to MCP servers"""
    
    def __init__(self, cli_path: str = "manus-mcp-cli", timeout: int = 30):
        """
        Initialize the MCP connector
        
        Args:
            cli_path: Path to manus-mcp-cli executable
            timeout: Command timeout in seconds
        """
        self.cli_path = cli_path
        self.timeout = timeout
    
    def _execute_command(self, args: List[str]) -> Dict[str, Any]:
        """
        Execute a manus-mcp-cli command
        
        Args:
            args: Command arguments
            
        Returns:
            Dictionary containing command output or error information
        """
        try:
            result = subprocess.run(
                [self.cli_path] + args,
                capture_output=True,
                text=True,
                timeout=self.timeout
            )
            
            if result.returncode == 0:
                # Try to parse as JSON
                try:
                    return {
                        "success": True,
                        "data": json.loads(result.stdout) if result.stdout.strip() else {},
                        "raw_output": result.stdout
                    }
                except json.JSONDecodeError:
                    # Return raw output if not JSON
                    return {
                        "success": True,
                        "data": {},
                        "raw_output": result.stdout
                    }
            else:
                return {
                    "success": False,
                    "error": result.stderr or result.stdout,
                    "returncode": result.returncode
                }
                
        except subprocess.TimeoutExpired:
            return {
                "success": False,
                "error": f"Command timed out after {self.timeout} seconds"
            }
        except FileNotFoundError:
            return {
                "success": False,
                "error": f"manus-mcp-cli not found at {self.cli_path}"
            }
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def discover_servers(self) -> List[MCPServer]:
        """
        Discover available MCP servers
        
        Returns:
            List of MCPServer objects
        """
        result = self._execute_command(["server", "list"])
        
        if not result["success"]:
            return []
        
        servers = []
        raw_output = result.get("raw_output", "")
        
        # Check if no servers found
        if "No MCP servers found" in raw_output:
            return []
        
        # Parse server list from output
        # The output format may vary, so we handle both JSON and text formats
        if result["data"]:
            # JSON format
            if isinstance(result["data"], list):
                for server_data in result["data"]:
                    if isinstance(server_data, dict):
                        servers.append(MCPServer(
                            name=server_data.get("name", ""),
                            status="Unknown"
                        ))
                    elif isinstance(server_data, str):
                        servers.append(MCPServer(name=server_data, status="Unknown"))
            elif isinstance(result["data"], dict):
                # Single server or server dict
                for name in result["data"].keys():
                    servers.append(MCPServer(name=name, status="Unknown"))
        else:
            # Text format - parse lines
            lines = raw_output.strip().split('\n')
            for line in lines:
                line = line.strip()
                if line and not line.startswith('#') and not line.startswith('No '):
                    # Extract server name (may have additional info)
                    parts = line.split()
                    if parts:
                        servers.append(MCPServer(name=parts[0], status="Unknown"))
        
        return servers
    
    def check_server_connection(self, server_name: str) -> MCPServer:
        """
        Check connection status of an MCP server
        
        Args:
            server_name: Name of the server to check
            
        Returns:
            MCPServer object with updated status
        """
        result = self._execute_command(["server", "check", "-s", server_name])
        
        if result["success"]:
            return MCPServer(name=server_name, status="Connected")
        else:
            return MCPServer(
                name=server_name,
                status="Disconnected",
                error_message=result.get("error", "Unknown error")
            )
    
    def list_tools(self, server_name: str) -> List[MCPCapability]:
        """
        List all tools available on an MCP server
        
        Args:
            server_name: Name of the server
            
        Returns:
            List of MCPCapability objects representing tools
        """
        result = self._execute_command(["tool", "list", "-s", server_name])
        
        if not result["success"]:
            return []
        
        tools = []
        
        # Parse JSON response
        if result["data"]:
            if isinstance(result["data"], list):
                for tool_data in result["data"]:
                    if isinstance(tool_data, dict):
                        tools.append(MCPCapability(
                            name=tool_data.get("name", ""),
                            description=tool_data.get("description"),
                            capability_type="tool",
                            details=tool_data
                        ))
                    elif isinstance(tool_data, str):
                        tools.append(MCPCapability(
                            name=tool_data,
                            capability_type="tool"
                        ))
            elif isinstance(result["data"], dict):
                # Handle dict format
                if "tools" in result["data"]:
                    for tool_data in result["data"]["tools"]:
                        tools.append(MCPCapability(
                            name=tool_data.get("name", ""),
                            description=tool_data.get("description"),
                            capability_type="tool",
                            details=tool_data
                        ))
        
        return tools
    
    def list_resources(self, server_name: str) -> List[MCPCapability]:
        """
        List all resources available on an MCP server
        
        Args:
            server_name: Name of the server
            
        Returns:
            List of MCPCapability objects representing resources
        """
        result = self._execute_command(["resource", "list", "-s", server_name])
        
        if not result["success"]:
            return []
        
        resources = []
        
        # Parse JSON response
        if result["data"]:
            if isinstance(result["data"], list):
                for resource_data in result["data"]:
                    if isinstance(resource_data, dict):
                        resources.append(MCPCapability(
                            name=resource_data.get("uri") or resource_data.get("name", ""),
                            description=resource_data.get("description"),
                            capability_type="resource",
                            details=resource_data
                        ))
                    elif isinstance(resource_data, str):
                        resources.append(MCPCapability(
                            name=resource_data,
                            capability_type="resource"
                        ))
            elif isinstance(result["data"], dict):
                # Handle dict format
                if "resources" in result["data"]:
                    for resource_data in result["data"]["resources"]:
                        resources.append(MCPCapability(
                            name=resource_data.get("uri") or resource_data.get("name", ""),
                            description=resource_data.get("description"),
                            capability_type="resource",
                            details=resource_data
                        ))
        
        return resources
    
    def list_prompts(self, server_name: str) -> List[MCPCapability]:
        """
        List all prompts available on an MCP server
        
        Args:
            server_name: Name of the server
            
        Returns:
            List of MCPCapability objects representing prompts
        """
        result = self._execute_command(["prompt", "list", "-s", server_name])
        
        if not result["success"]:
            return []
        
        prompts = []
        
        # Parse JSON response
        if result["data"]:
            if isinstance(result["data"], list):
                for prompt_data in result["data"]:
                    if isinstance(prompt_data, dict):
                        prompts.append(MCPCapability(
                            name=prompt_data.get("name", ""),
                            description=prompt_data.get("description"),
                            capability_type="prompt",
                            details=prompt_data
                        ))
                    elif isinstance(prompt_data, str):
                        prompts.append(MCPCapability(
                            name=prompt_data,
                            capability_type="prompt"
                        ))
            elif isinstance(result["data"], dict):
                # Handle dict format
                if "prompts" in result["data"]:
                    for prompt_data in result["data"]["prompts"]:
                        prompts.append(MCPCapability(
                            name=prompt_data.get("name", ""),
                            description=prompt_data.get("description"),
                            capability_type="prompt",
                            details=prompt_data
                        ))
        
        return prompts
    
    def get_all_capabilities(self, server_name: str) -> Dict[str, List[MCPCapability]]:
        """
        Get all capabilities (tools, resources, prompts) for a server
        
        Args:
            server_name: Name of the server
            
        Returns:
            Dictionary with keys 'tools', 'resources', 'prompts'
        """
        return {
            "tools": self.list_tools(server_name),
            "resources": self.list_resources(server_name),
            "prompts": self.list_prompts(server_name)
        }
    
    def get_tool_details(self, server_name: str, tool_name: str) -> Optional[Dict[str, Any]]:
        """
        Get detailed information about a specific tool
        
        Args:
            server_name: Name of the server
            tool_name: Name of the tool
            
        Returns:
            Dictionary with tool details or None if error
        """
        result = self._execute_command(["tool", "get", tool_name, "-s", server_name])
        return result["data"] if result["success"] else None
    
    def get_resource_details(self, server_name: str, resource_uri: str) -> Optional[Dict[str, Any]]:
        """
        Get detailed information about a specific resource
        
        Args:
            server_name: Name of the server
            resource_uri: URI of the resource
            
        Returns:
            Dictionary with resource details or None if error
        """
        result = self._execute_command(["resource", "get", resource_uri, "-s", server_name])
        return result["data"] if result["success"] else None
    
    def get_prompt_details(self, server_name: str, prompt_name: str) -> Optional[Dict[str, Any]]:
        """
        Get detailed information about a specific prompt
        
        Args:
            server_name: Name of the server
            prompt_name: Name of the prompt
            
        Returns:
            Dictionary with prompt details or None if error
        """
        result = self._execute_command(["prompt", "get", prompt_name, "-s", server_name])
        return result["data"] if result["success"] else None
    
    def call_tool(self, server_name: str, tool_name: str, arguments: Dict[str, Any]) -> Dict[str, Any]:
        """
        Execute a tool on an MCP server
        
        Args:
            server_name: Name of the server
            tool_name: Name of the tool
            arguments: Tool arguments as dictionary
            
        Returns:
            Dictionary with execution result
        """
        # Convert arguments to JSON string
        args_json = json.dumps(arguments)
        
        result = self._execute_command([
            "tool", "call", tool_name,
            "-s", server_name,
            "--input", args_json
        ])
        
        return result
    
    def read_resource(self, server_name: str, resource_uri: str) -> Dict[str, Any]:
        """
        Read a resource from an MCP server
        
        Args:
            server_name: Name of the server
            resource_uri: URI of the resource
            
        Returns:
            Dictionary with resource content
        """
        result = self._execute_command(["resource", "read", resource_uri, "-s", server_name])
        return result
    
    def call_prompt(self, server_name: str, prompt_name: str, arguments: Dict[str, Any]) -> Dict[str, Any]:
        """
        Call a prompt on an MCP server
        
        Args:
            server_name: Name of the server
            prompt_name: Name of the prompt
            arguments: Prompt arguments as dictionary
            
        Returns:
            Dictionary with prompt result
        """
        # Convert arguments to JSON string
        args_json = json.dumps(arguments)
        
        result = self._execute_command([
            "prompt", "call", prompt_name,
            "-s", server_name,
            "--arguments", args_json
        ])
        
        return result
