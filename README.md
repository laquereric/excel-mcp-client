> # Excel MCP Plugin

This repository contains a Python-based Excel plugin that enables seamless connection to Model Context Protocol (MCP) servers. The plugin provides a user-friendly interface within Excel to discover, monitor, and interact with MCP servers and their capabilities, including tools, resources, and prompts.

## Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Architecture](#architecture)
- [Getting Started](#getting-started)
- [Usage](#usage)
- [Configuration](#configuration)
- [Security](#security)
- [Contributing](#contributing)
- [License](#license)

## Overview

The Excel MCP Plugin bridges the gap between the powerful capabilities of MCP servers and the familiar environment of Microsoft Excel. It allows users to leverage AI-driven tools and contextual data directly within their spreadsheets, enhancing productivity and enabling complex workflows without leaving Excel.

The plugin dynamically generates a dashboard in a dedicated "MCPs" sheet, providing a real-time overview of all available MCP servers and their status. It organizes each server's capabilities—tools, resources, and prompts—into a clear and structured layout, making it easy to see what each server can do.

## Features

- **MCP Server Discovery**: Automatically discovers all available MCP servers on the network.
- **Real-time Status Monitoring**: Displays the connection status of each MCP server ("Connected," "Disconnected," or "Error").
- **Capability Dashboard**: Organizes tools, resources, and prompts for each server in a structured and easy-to-read format.
- **Interactive Interface**: Provides ribbon buttons and user-defined functions (UDFs) for seamless interaction.
- **Tool Execution**: Execute MCP tools directly from Excel using the `=MCP_CALL_TOOL()` function.
- **Resource Access**: Read content from MCP resources using the `=MCP_READ_RESOURCE()` function.
- **Prompt Invocation**: Utilize MCP prompts with dynamic arguments using the `=MCP_GET_PROMPT()` function.
- **Detailed Capability View**: Get detailed information about any tool, resource, or prompt.

## Architecture

The plugin is built on a modular architecture that separates concerns and allows for easy extension and maintenance.

| Component                 | Description                                                                                               |
| ------------------------- | --------------------------------------------------------------------------------------------------------- |
| **Excel Add-in Interface**  | Built with `xlwings`, this component provides the user interface, including ribbon buttons and UDFs.          |
| **MCP Connection Manager**  | A Python module (`mcp_connector.py`) that interfaces with the `manus-mcp-cli` to communicate with MCP servers. |
| **Data Synchronization Engine** | The `sheet_manager.py` module is responsible for updating the "MCPs" sheet with the latest server data.    |

## Getting Started

### Prerequisites

- Python 3.7+
- Microsoft Excel (2007 or later on Windows, 2016 or later on macOS)
- `manus-mcp-cli` installed and configured

### Installation

1.  **Clone the repository:**

    ```bash
    git clone https://github.com/yourusername/excel-mcp-plugin.git
    cd excel-mcp-plugin
    ```

2.  **Install dependencies:**

    ```bash
    pip install -r requirements.txt
    ```

3.  **Install the xlwings add-in:**

    ```bash
    xlwings addin install
    ```

## Usage

### Initializing the Plugin

1.  Open a new or existing Excel workbook.
2.  A new "MCP Plugin" tab should appear in the Excel ribbon.
3.  Click the **Initialize Sheet** button to create the "MCPs" dashboard.

### Refreshing Server Data

- Click the **Refresh MCPs** button in the ribbon to update the sheet with the latest server information.

### Using User-Defined Functions (UDFs)

The plugin provides several UDFs to interact with MCP servers directly from Excel cells:

- `=MCP_CALL_TOOL(server, tool, arguments)`: Executes an MCP tool.
- `=MCP_READ_RESOURCE(server, resource_uri)`: Reads an MCP resource.
- `=MCP_GET_PROMPT(server, prompt, arguments)`: Gets a formatted prompt string.
- `=MCP_SERVER_STATUS(server)`: Checks the status of a specific server.
- `=MCP_LIST_SERVERS()`: Lists all discovered MCP servers.

## Configuration

The plugin can be configured via the `mcp_plugin_config.json` file located in the same directory as the plugin.

| Setting                     | Description                                                                 |
| --------------------------- | --------------------------------------------------------------------------- |
| `auto_refresh`              | Enable or disable automatic refreshing of the MCPs sheet.                   |
| `refresh_interval_seconds`  | The interval in seconds for automatic refreshing.                           |
| `show_disconnected_servers` | Whether to display servers that are currently disconnected.                 |
| `cli_timeout_seconds`       | The timeout in seconds for `manus-mcp-cli` commands.                        |
| `cli_path`                  | The file path to the `manus-mcp-cli` executable.                            |

## Security

The plugin is designed with security in mind:

- **User Confirmation**: All tool executions require explicit user confirmation.
- **No Automatic Execution**: Tools are never executed without direct user interaction.
- **Secure Credential Storage**: Server credentials (if any) should be managed by the `manus-mcp-cli` and are not stored by the plugin.

## Contributing

Contributions are welcome! Please feel free to submit a pull request or open an issue to discuss any changes.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

