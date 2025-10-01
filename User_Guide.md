> # Excel MCP Plugin - User Guide

## 1. Introduction

Welcome to the Excel MCP Plugin! This guide provides everything you need to know to install, configure, and use the plugin to connect Microsoft Excel with Model Context Protocol (MCP) servers. This powerful integration allows you to leverage AI-driven tools, access real-time data resources, and utilize standardized prompts directly within your spreadsheets.

## 2. Features

- **Dynamic MCP Dashboard**: Automatically generates and maintains a sheet named "MCPs" that provides a comprehensive overview of all discovered servers.
- **Real-time Status**: See the live connection status of each MCP server.
- **Structured Capability View**: Tools, resources, and prompts for each server are neatly organized into collapsible row groups.
- **Interactive Ribbon Menu**: A dedicated "MCP Plugin" tab in the Excel ribbon for easy access to key functions.
- **Powerful UDFs**: A suite of User-Defined Functions (e.g., `=MCP_CALL_TOOL`) to interact with MCP servers from any cell in your workbook.

## 3. Installation

### Prerequisites

- **Microsoft Excel**: Windows (2007 or later) or macOS (2016 or later).
- **Python**: Version 3.7 or higher.
- **`manus-mcp-cli`**: The MCP command-line interface must be installed and accessible in your system's PATH.

### Step-by-Step Installation

1.  **Download the Plugin**: Download the `excel-mcp-plugin.zip` file and extract it to a permanent location on your computer.

2.  **Install Python Dependencies**: Open a terminal or command prompt, navigate to the extracted folder, and run:

    ```bash
    pip install -r requirements.txt
    ```

3.  **Install the Excel Add-in**: In the same terminal, run the following `xlwings` command to install the add-in:

    ```bash
    xlwings addin install
    ```

    This command registers the plugin with Excel. The next time you open Excel, the "MCP Plugin" tab will be available in the ribbon.

## 4. The MCPs Dashboard

The core of the plugin is the "MCPs" sheet, which acts as your central dashboard for all MCP-related activities.

### Initializing the Dashboard

- To create the dashboard, click the **Initialize Sheet** button in the "MCP Plugin" ribbon tab. This will create the "MCPs" sheet with the required structure.

### Understanding the Layout

- **Columns**: Each MCP server discovered on your network is assigned a column.
- **Rows**: Capabilities are organized in rows. The first column, "Capability Type," lists all unique tools, resources, and prompts available across all servers.
- **Row Groups**: The sheet is divided into three main sections: **TOOLS**, **RESOURCES**, and **PROMPTS**. 
- **Status**: The "Status" row indicates whether each server is "Connected," "Disconnected," or has encountered an "Error."
- **Checkmarks (✓)**: A checkmark indicates that a specific server possesses the capability listed in that row.

## 5. Using the Ribbon Menu

The "MCP Plugin" ribbon provides quick access to the main functions:

- **Refresh MCPs**: Updates the entire "MCPs" sheet. It re-discovers servers, checks their status, and updates all capability listings.
- **Initialize Sheet**: Creates the "MCPs" sheet if it doesn't exist.
- **View Details**: Select any capability cell (a checkmark) and click this button to see detailed information about that tool, resource, or prompt in a pop-up window.
- **Clear Sheet**: Resets the "MCPs" sheet to its initial empty state.

## 6. User-Defined Functions (UDFs)

The plugin's UDFs allow you to pull MCP data directly into any cell. You can use them just like any other Excel formula.

### `=MCP_CALL_TOOL(server, tool, [arguments])`

Executes a tool on a specified MCP server.

- **server**: The name of the MCP server (e.g., `"weather-mcp"`).
- **tool**: The name of the tool to execute (e.g., `"get_current_weather"`).
- **arguments** (optional): A JSON string of the arguments for the tool (e.g., `"{\"city\":\"London\"}"`).

**Example:**
`=MCP_CALL_TOOL("weather-mcp", "get_current_weather", "{\"city\":\"New York\"}")`

### `=MCP_READ_RESOURCE(server, resource_uri)`

Reads the content of a resource from an MCP server.

- **server**: The name of the MCP server.
- **resource_uri**: The URI of the resource to read (e.g., `"db://schema"`).

**Example:**
`=MCP_READ_RESOURCE("database-mcp", "db://tables")`

### `=MCP_GET_PROMPT(server, prompt, [arguments])`

Retrieves a formatted prompt string from an MCP server.

- **server**: The name of the MCP server.
- **prompt**: The name of the prompt.
- **arguments** (optional): A JSON string of arguments to fill into the prompt template.

**Example:**
`=MCP_GET_PROMPT("summary-mcp", "summarize_text", "{\"text\":\"A long article...\"}")`

### Other Helper Functions

- `=MCP_SERVER_STATUS(server)`: Returns the connection status of a single server.
- `=MCP_LIST_SERVERS()`: Returns a comma-separated string of all discovered server names.

## 7. Configuration

You can fine-tune the plugin's behavior by editing the `mcp_plugin_config.json` file.

- **`auto_refresh`**: Set to `true` to have the sheet refresh automatically at a set interval.
- **`refresh_interval_seconds`**: The time in seconds between automatic refreshes.
- **`cli_path`**: If `manus-mcp-cli` is not in your system's PATH, you can specify the full path to the executable here.

## 8. Troubleshooting

- **"#NAME?" Error**: This usually means the xlwings add-in is not properly installed or enabled. Try running `xlwings addin install` again.
- **No Servers Found**: Ensure that `manus-mcp-cli` is correctly configured and can discover servers from your command line. The plugin relies entirely on the CLI for discovery.
- **UDFs Not Calculating**: Make sure the Python environment where you installed the dependencies is running. The xlwings add-in will show the status of the Python connection.

--- 
*Thank you for using the Excel MCP Plugin!*
