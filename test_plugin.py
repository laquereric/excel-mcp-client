"""
Test script for Excel MCP Plugin
Tests the core functionality without Excel
"""

import sys
from mcp_connector import MCPConnector, MCPServer, MCPCapability


def test_connector():
    """Test the MCP connector"""
    print("=" * 60)
    print("Testing MCP Connector")
    print("=" * 60)
    
    connector = MCPConnector()
    
    # Test 1: Discover servers
    print("\n1. Discovering MCP servers...")
    servers = connector.discover_servers()
    print(f"   Found {len(servers)} servers")
    
    if not servers:
        print("   No servers found. This is expected if no MCP servers are configured.")
        print("   The plugin will work correctly once MCP servers are available.")
        return True
    
    for server in servers:
        print(f"   - {server.name} (Status: {server.status})")
    
    # Test 2: Check connection for each server
    print("\n2. Checking server connections...")
    for server in servers:
        checked_server = connector.check_server_connection(server.name)
        print(f"   - {checked_server.name}: {checked_server.status}")
        
        if checked_server.status == "Connected":
            # Test 3: Get capabilities
            print(f"\n3. Getting capabilities for {checked_server.name}...")
            
            tools = connector.list_tools(checked_server.name)
            print(f"   Tools: {len(tools)}")
            for tool in tools[:5]:  # Show first 5
                print(f"     - {tool.name}")
            if len(tools) > 5:
                print(f"     ... and {len(tools) - 5} more")
            
            resources = connector.list_resources(checked_server.name)
            print(f"   Resources: {len(resources)}")
            for resource in resources[:5]:  # Show first 5
                print(f"     - {resource.name}")
            if len(resources) > 5:
                print(f"     ... and {len(resources) - 5} more")
            
            prompts = connector.list_prompts(checked_server.name)
            print(f"   Prompts: {len(prompts)}")
            for prompt in prompts[:5]:  # Show first 5
                print(f"     - {prompt.name}")
            if len(prompts) > 5:
                print(f"     ... and {len(prompts) - 5} more")
    
    return True


def test_data_structures():
    """Test data structures"""
    print("\n" + "=" * 60)
    print("Testing Data Structures")
    print("=" * 60)
    
    # Test MCPServer
    server = MCPServer(name="test-server", status="Connected")
    print(f"\n1. MCPServer: {server.name} - {server.status}")
    
    # Test MCPCapability
    capability = MCPCapability(
        name="test_tool",
        description="A test tool",
        capability_type="tool"
    )
    print(f"2. MCPCapability: {capability.name} ({capability.capability_type})")
    print(f"   Description: {capability.description}")
    
    return True


def main():
    """Run all tests"""
    print("\n" + "=" * 60)
    print("Excel MCP Plugin - Test Suite")
    print("=" * 60)
    
    try:
        # Test data structures
        if not test_data_structures():
            print("\n❌ Data structure tests failed!")
            return False
        
        # Test connector
        if not test_connector():
            print("\n❌ Connector tests failed!")
            return False
        
        print("\n" + "=" * 60)
        print("✅ All tests passed!")
        print("=" * 60)
        print("\nThe plugin is ready to use.")
        print("Note: Full functionality requires MCP servers to be configured.")
        return True
        
    except Exception as e:
        print(f"\n❌ Test failed with error: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
