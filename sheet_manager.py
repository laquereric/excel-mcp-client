"""
Sheet Manager Module
Manages the MCPs sheet structure and updates
"""

import xlwings as xw
from typing import List, Dict, Optional, Tuple
from mcp_connector import MCPConnector, MCPServer, MCPCapability


class SheetManager:
    """Manages the MCPs sheet in Excel workbooks"""
    
    SHEET_NAME = "MCPs"
    HEADER_ROW = 1
    STATUS_ROW = 2
    FIRST_DATA_ROW = 4
    
    # Colors (RGB)
    COLOR_HEADER = (68, 114, 196)  # Blue
    COLOR_SECTION = (217, 217, 217)  # Light gray
    COLOR_CONNECTED = (198, 224, 180)  # Light green
    COLOR_DISCONNECTED = (255, 199, 206)  # Light red
    COLOR_ERROR = (255, 235, 156)  # Light yellow
    
    def __init__(self, workbook: xw.Book, connector: MCPConnector):
        """
        Initialize the sheet manager
        
        Args:
            workbook: xlwings workbook object
            connector: MCPConnector instance
        """
        self.workbook = workbook
        self.connector = connector
    
    def ensure_sheet_exists(self) -> xw.Sheet:
        """
        Ensure the MCPs sheet exists, create if not
        
        Returns:
            xlwings Sheet object
        """
        try:
            sheet = self.workbook.sheets[self.SHEET_NAME]
        except KeyError:
            # Sheet doesn't exist, create it
            sheet = self.workbook.sheets.add(self.SHEET_NAME)
            self._initialize_sheet_structure(sheet)
        
        return sheet
    
    def _initialize_sheet_structure(self, sheet: xw.Sheet):
        """
        Initialize the basic structure of the MCPs sheet
        
        Args:
            sheet: xlwings Sheet object
        """
        # Set up header row
        sheet.range("A1").value = "Capability Type"
        sheet.range("A1").color = self.COLOR_HEADER
        sheet.range("A1").api.Font.Bold = True
        sheet.range("A1").api.Font.Color = 0xFFFFFF  # White text
        
        # Set up status row
        sheet.range("A2").value = "Status"
        sheet.range("A2").api.Font.Bold = True
        
        # Set column width
        sheet.range("A:A").column_width = 20
        
        # Freeze first column and first two rows
        sheet.range("B3").select()
        self.workbook.api.ActiveWindow.FreezePanes = True
    
    def _get_next_column(self, sheet: xw.Sheet) -> int:
        """
        Get the next available column for a new server
        
        Args:
            sheet: xlwings Sheet object
            
        Returns:
            Column number (1-based)
        """
        # Find the last used column in the header row
        last_col = 1
        col = 2  # Start from column B
        
        while sheet.range((self.HEADER_ROW, col)).value is not None:
            last_col = col
            col += 1
        
        return last_col + 1 if last_col > 1 else 2
    
    def _find_server_column(self, sheet: xw.Sheet, server_name: str) -> Optional[int]:
        """
        Find the column number for a specific server
        
        Args:
            sheet: xlwings Sheet object
            server_name: Name of the server to find
            
        Returns:
            Column number (1-based) or None if not found
        """
        col = 2  # Start from column B
        
        while True:
            value = sheet.range((self.HEADER_ROW, col)).value
            if value is None:
                return None
            if value == server_name:
                return col
            col += 1
    
    def _clear_column(self, sheet: xw.Sheet, col: int):
        """
        Clear all data in a column (except header)
        
        Args:
            sheet: xlwings Sheet object
            col: Column number to clear
        """
        # Clear from status row downward
        last_row = sheet.cells.last_cell.row
        if last_row >= self.STATUS_ROW:
            sheet.range((self.STATUS_ROW, col), (last_row, col)).clear_contents()
            sheet.range((self.STATUS_ROW, col), (last_row, col)).color = None
    
    def _write_server_column(self, sheet: xw.Sheet, col: int, server: MCPServer, 
                            capabilities: Dict[str, List[MCPCapability]]):
        """
        Write server data to a column
        
        Args:
            sheet: xlwings Sheet object
            col: Column number to write to
            server: MCPServer object
            capabilities: Dictionary of capabilities
        """
        # Write server name in header
        header_cell = sheet.range((self.HEADER_ROW, col))
        header_cell.value = server.name
        header_cell.color = self.COLOR_HEADER
        header_cell.api.Font.Bold = True
        header_cell.api.Font.Color = 0xFFFFFF  # White text
        
        # Write status
        status_cell = sheet.range((self.STATUS_ROW, col))
        status_cell.value = server.status
        
        # Color code status
        if server.status == "Connected":
            status_cell.color = self.COLOR_CONNECTED
        elif server.status == "Disconnected":
            status_cell.color = self.COLOR_DISCONNECTED
        else:
            status_cell.color = self.COLOR_ERROR
        
        # Set column width
        sheet.range((1, col), (1, col)).column_width = 18
        
        # Write capabilities
        current_row = self.FIRST_DATA_ROW
        
        # Write TOOLS section
        current_row = self._write_capability_section(
            sheet, col, current_row, "TOOLS", capabilities.get("tools", [])
        )
        
        # Add separator
        current_row += 1
        
        # Write RESOURCES section
        current_row = self._write_capability_section(
            sheet, col, current_row, "RESOURCES", capabilities.get("resources", [])
        )
        
        # Add separator
        current_row += 1
        
        # Write PROMPTS section
        current_row = self._write_capability_section(
            sheet, col, current_row, "PROMPTS", capabilities.get("prompts", [])
        )
    
    def _write_capability_section(self, sheet: xw.Sheet, col: int, start_row: int,
                                  section_name: str, capabilities: List[MCPCapability]) -> int:
        """
        Write a capability section (TOOLS, RESOURCES, or PROMPTS)
        
        Args:
            sheet: xlwings Sheet object
            col: Column number
            start_row: Starting row number
            section_name: Name of the section
            capabilities: List of capabilities
            
        Returns:
            Next available row number
        """
        current_row = start_row
        
        # Write section header in column A if not already present
        if sheet.range((current_row, 1)).value != section_name:
            section_cell = sheet.range((current_row, 1))
            section_cell.value = section_name
            section_cell.api.Font.Bold = True
            section_cell.color = self.COLOR_SECTION
        
        # Write checkmark or capability indicator in the server column
        if capabilities:
            # Mark that this server has this capability type
            sheet.range((current_row, col)).value = "✓"
            sheet.range((current_row, col)).api.Font.Size = 14
            current_row += 1
            
            # Write individual capabilities
            for capability in capabilities:
                # Write capability name in column A if not present
                if sheet.range((current_row, 1)).value != capability.name:
                    sheet.range((current_row, 1)).value = capability.name
                
                # Mark this capability as available for this server
                sheet.range((current_row, col)).value = "✓"
                current_row += 1
        else:
            # No capabilities in this section
            sheet.range((current_row, col)).value = "-"
            current_row += 1
        
        return current_row
    
    def update_sheet(self) -> Tuple[int, int, List[str]]:
        """
        Update the MCPs sheet with current server data
        
        Returns:
            Tuple of (servers_found, servers_connected, error_messages)
        """
        sheet = self.ensure_sheet_exists()
        
        # Discover servers
        servers = self.connector.discover_servers()
        
        if not servers:
            return (0, 0, ["No MCP servers found"])
        
        connected_count = 0
        errors = []
        
        # Process each server
        for server in servers:
            try:
                # Check connection
                server = self.connector.check_server_connection(server.name)
                
                if server.status == "Connected":
                    connected_count += 1
                    
                    # Get capabilities
                    capabilities = self.connector.get_all_capabilities(server.name)
                    
                    # Find or create column for this server
                    col = self._find_server_column(sheet, server.name)
                    if col is None:
                        col = self._get_next_column(sheet)
                    else:
                        # Clear existing data
                        self._clear_column(sheet, col)
                    
                    # Write server data
                    self._write_server_column(sheet, col, server, capabilities)
                else:
                    # Server not connected, still show it
                    col = self._find_server_column(sheet, server.name)
                    if col is None:
                        col = self._get_next_column(sheet)
                    else:
                        self._clear_column(sheet, col)
                    
                    # Write header and status only
                    header_cell = sheet.range((self.HEADER_ROW, col))
                    header_cell.value = server.name
                    header_cell.color = self.COLOR_HEADER
                    header_cell.api.Font.Bold = True
                    header_cell.api.Font.Color = 0xFFFFFF
                    
                    status_cell = sheet.range((self.STATUS_ROW, col))
                    status_cell.value = server.status
                    status_cell.color = self.COLOR_DISCONNECTED
                    
                    errors.append(f"{server.name}: {server.error_message or 'Connection failed'}")
                    
            except Exception as e:
                errors.append(f"{server.name}: {str(e)}")
        
        # Auto-fit rows
        sheet.autofit('r')
        
        return (len(servers), connected_count, errors)
    
    def get_capability_at_cell(self, sheet: xw.Sheet, row: int, col: int) -> Optional[Tuple[str, str, str]]:
        """
        Get capability information at a specific cell
        
        Args:
            sheet: xlwings Sheet object
            row: Row number
            col: Column number
            
        Returns:
            Tuple of (server_name, capability_type, capability_name) or None
        """
        # Get server name from header
        server_name = sheet.range((self.HEADER_ROW, col)).value
        if not server_name:
            return None
        
        # Get capability name from column A
        capability_name = sheet.range((row, 1)).value
        if not capability_name:
            return None
        
        # Determine capability type by searching upward for section header
        search_row = row
        while search_row >= self.FIRST_DATA_ROW:
            value = sheet.range((search_row, 1)).value
            if value in ["TOOLS", "RESOURCES", "PROMPTS"]:
                capability_type = value.lower().rstrip('s')  # Remove trailing 's'
                return (server_name, capability_type, capability_name)
            search_row -= 1
        
        return None
    
    def clear_sheet(self):
        """Clear all data from the MCPs sheet"""
        try:
            sheet = self.workbook.sheets[self.SHEET_NAME]
            sheet.clear()
            self._initialize_sheet_structure(sheet)
        except KeyError:
            pass  # Sheet doesn't exist, nothing to clear
