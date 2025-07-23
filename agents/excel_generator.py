import pandas as pd
from datetime import datetime
from typing import Dict, List, Any
import xlsxwriter

class ESGFormGenerator:
    """Generate comprehensive Excel forms for ESG data capture."""
    
    def create_comprehensive_form(self, metrics_data: Dict[str, Any], 
                                properties: List[Dict[str, str]], 
                                investor_name: str,
                                output_path: str):
        """
        Create a comprehensive ESG data capture form in Excel.
        
        Args:
            metrics_data: Structured ESG metrics from crew analysis
            properties: List of properties to track
            investor_name: Name of the investor
            output_path: Path to save the Excel file
        """
        # Create Excel workbook
        workbook = xlsxwriter.Workbook(output_path)
        
        # Define formats
        formats = self._create_formats(workbook)
        
        # Create sheets
        self._create_overview_sheet(workbook, formats, investor_name, properties)
        self._create_metrics_summary_sheet(workbook, formats, metrics_data)
        self._create_data_entry_sheet(workbook, formats, metrics_data, properties)
        
        # Create individual property sheets
        for prop in properties:
            self._create_property_sheet(workbook, formats, metrics_data, prop)
        
        # Add data validation sheet
        self._create_validation_sheet(workbook, formats)
        
        # Close workbook
        workbook.close()
    
    def _create_formats(self, workbook) -> Dict[str, Any]:
        """Create all formatting styles for the workbook."""
        return {
            'title': workbook.add_format({
                'bold': True,
                'font_size': 16,
                'font_color': 'white',
                'bg_color': '#1f4788',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            }),
            'header': workbook.add_format({
                'bold': True,
                'font_size': 12,
                'font_color': 'white',
                'bg_color': '#4472c4',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'text_wrap': True
            }),
            'category': workbook.add_format({
                'bold': True,
                'font_size': 11,
                'bg_color': '#d9e2f3',
                'border': 1,
                'indent': 1
            }),
            'subcategory': workbook.add_format({
                'bold': True,
                'font_size': 10,
                'bg_color': '#e7eef8',
                'border': 1,
                'indent': 2
            }),
            'metric': workbook.add_format({
                'font_size': 10,
                'border': 1,
                'indent': 3,
                'text_wrap': True
            }),
            'input_numeric': workbook.add_format({
                'font_size': 10,
                'border': 1,
                'bg_color': '#fff2cc',
                'num_format': '#,##0.00'
            }),
            'input_text': workbook.add_format({
                'font_size': 10,
                'border': 1,
                'bg_color': '#fff2cc'
            }),
            'input_percent': workbook.add_format({
                'font_size': 10,
                'border': 1,
                'bg_color': '#fff2cc',
                'num_format': '0.0%'
            }),
            'input_currency': workbook.add_format({
                'font_size': 10,
                'border': 1,
                'bg_color': '#fff2cc',
                'num_format': '$#,##0.00'
            }),
            'note': workbook.add_format({
                'font_size': 9,
                'font_color': '#7f7f7f',
                'italic': True
            }),
            'timestamp': workbook.add_format({
                'font_size': 9,
                'font_color': '#7f7f7f',
                'num_format': 'yyyy-mm-dd hh:mm:ss'
            })
        }
    
    def _create_overview_sheet(self, workbook, formats: Dict, investor_name: str, properties: List[Dict]):
        """Create the overview sheet with instructions."""
        worksheet = workbook.add_worksheet('Overview')
        
        # Set column widths
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:E', 15)
        
        # Title
        worksheet.merge_range('A1:E1', f'ESG Data Capture Form - {investor_name}', formats['title'])
        worksheet.set_row(0, 30)
        
        # Generation info
        row = 2
        worksheet.write(row, 0, 'Generated:', formats['metric'])
        worksheet.write(row, 1, datetime.now(), formats['timestamp'])
        
        row += 1
        worksheet.write(row, 0, 'Total Properties:', formats['metric'])
        worksheet.write(row, 1, len(properties))
        
        # Instructions
        row += 2
        worksheet.merge_range(f'A{row+1}:E{row+1}', 'Instructions', formats['header'])
        
        instructions = [
            "1. Yellow cells indicate required input fields",
            "2. Complete data for each property in its respective sheet",
            "3. Use the 'Data Entry' sheet for consolidated reporting",
            "4. Ensure all units match the specified requirements",
            "5. Save the file regularly to prevent data loss",
            "6. Review the 'Validation Rules' sheet for data requirements"
        ]
        
        row += 2
        for instruction in instructions:
            worksheet.write(row, 0, instruction, formats['metric'])
            row += 1
        
        # Property list
        row += 2
        worksheet.merge_range(f'A{row+1}:C{row+1}', 'Properties Included', formats['header'])
        row += 2
        
        worksheet.write(row, 0, 'Property Name', formats['subcategory'])
        worksheet.write(row, 1, 'Type', formats['subcategory'])
        worksheet.write(row, 2, 'Location', formats['subcategory'])
        
        row += 1
        for prop in properties:
            worksheet.write(row, 0, prop.get('name', ''), formats['metric'])
            worksheet.write(row, 1, prop.get('type', ''), formats['metric'])
            worksheet.write(row, 2, prop.get('location', ''), formats['metric'])
            row += 1
    
    def _create_metrics_summary_sheet(self, workbook, formats: Dict, metrics_data: Dict):
        """Create a summary sheet of all metrics."""
        worksheet = workbook.add_worksheet('Metrics Summary')
        
        # Set column widths
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 25)
        worksheet.set_column('C:C', 40)
        worksheet.set_column('D:D', 15)
        worksheet.set_column('E:E', 15)
        worksheet.set_column('F:F', 20)
        
        # Headers
        headers = ['Category', 'Subcategory', 'Metric', 'Unit', 'Frequency', 'Data Type']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, formats['header'])
        
        # Freeze top row
        worksheet.freeze_panes(1, 0)
        
        # Write metrics
        row = 1
        for category, subcategories in metrics_data.items():
            worksheet.write(row, 0, category, formats['category'])
            row += 1
            
            for subcategory, metrics in subcategories.items():
                worksheet.write(row, 1, subcategory, formats['subcategory'])
                row += 1
                
                for metric in metrics:
                    worksheet.write(row, 2, metric.get('name', ''), formats['metric'])
                    worksheet.write(row, 3, metric.get('unit', ''), formats['metric'])
                    worksheet.write(row, 4, metric.get('frequency', 'Annual'), formats['metric'])
                    worksheet.write(row, 5, metric.get('data_type', 'text'), formats['metric'])
                    row += 1
    
    def _create_data_entry_sheet(self, workbook, formats: Dict, metrics_data: Dict, properties: List[Dict]):
        """Create the main data entry sheet with all properties."""
        worksheet = workbook.add_worksheet('Data Entry')
        
        # Set column widths
        worksheet.set_column('A:C', 20)
        worksheet.set_column('D:D', 40)
        worksheet.set_column('E:E', 15)
        
        # Dynamic columns for properties
        property_start_col = 5  # Column F
        for i in range(len(properties)):
            worksheet.set_column(property_start_col + i, property_start_col + i, 15)
        
        # Headers
        worksheet.write(0, 0, 'Category', formats['header'])
        worksheet.write(0, 1, 'Subcategory', formats['header'])
        worksheet.write(0, 2, 'Metric', formats['header'])
        worksheet.write(0, 3, 'Description', formats['header'])
        worksheet.write(0, 4, 'Unit', formats['header'])
        
        # Property headers
        for i, prop in enumerate(properties):
            worksheet.write(0, property_start_col + i, prop['name'], formats['header'])
        
        # Freeze panes
        worksheet.freeze_panes(1, 5)
        
        # Write metrics with input cells
        row = 1
        for category, subcategories in metrics_data.items():
            # Category row
            worksheet.merge_range(row, 0, row, 4, category, formats['category'])
            row += 1
            
            for subcategory, metrics in subcategories.items():
                # Subcategory row
                worksheet.merge_range(row, 1, row, 4, subcategory, formats['subcategory'])
                row += 1
                
                for metric in metrics:
                    worksheet.write(row, 2, metric.get('name', ''), formats['metric'])
                    worksheet.write(row, 3, metric.get('description', ''), formats['metric'])
                    worksheet.write(row, 4, metric.get('unit', ''), formats['metric'])
                    
                    # Add input cells for each property
                    data_type = metric.get('data_type', 'text')
                    for i in range(len(properties)):
                        cell_format = self._get_input_format(formats, data_type)
                        worksheet.write(row, property_start_col + i, '', cell_format)
                        
                        # Add validation if needed
                        self._add_validation(worksheet, row, property_start_col + i, data_type)
                    
                    row += 1
    
    def _create_property_sheet(self, workbook, formats: Dict, metrics_data: Dict, property_info: Dict):
        """Create individual sheet for each property."""
        sheet_name = property_info['name'][:31]  # Excel sheet name limit
        worksheet = workbook.add_worksheet(sheet_name)
        
        # Set column widths
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 25)
        worksheet.set_column('C:C', 40)
        worksheet.set_column('D:D', 15)
        worksheet.set_column('E:E', 20)
        worksheet.set_column('F:F', 50)
        
        # Property header
        worksheet.merge_range('A1:F1', f"ESG Data - {property_info['name']}", formats['title'])
        worksheet.set_row(0, 30)
        
        # Property info
        row = 2
        worksheet.write(row, 0, 'Property Type:', formats['metric'])
        worksheet.write(row, 1, property_info.get('type', ''), formats['input_text'])
        row += 1
        worksheet.write(row, 0, 'Location:', formats['metric'])
        worksheet.write(row, 1, property_info.get('location', ''), formats['input_text'])
        
        # Metrics headers
        row += 2
        headers = ['Category', 'Subcategory', 'Metric', 'Unit', 'Value', 'Notes']
        for col, header in enumerate(headers):
            worksheet.write(row, col, header, formats['header'])
        
        # Write metrics
        row += 1
        for category, subcategories in metrics_data.items():
            worksheet.write(row, 0, category, formats['category'])
            row += 1
            
            for subcategory, metrics in subcategories.items():
                worksheet.write(row, 1, subcategory, formats['subcategory'])
                row += 1
                
                for metric in metrics:
                    worksheet.write(row, 2, metric.get('name', ''), formats['metric'])
                    worksheet.write(row, 3, metric.get('unit', ''), formats['metric'])
                    
                    # Value input cell
                    data_type = metric.get('data_type', 'text')
                    cell_format = self._get_input_format(formats, data_type)
                    worksheet.write(row, 4, '', cell_format)
                    self._add_validation(worksheet, row, 4, data_type)
                    
                    # Notes cell
                    worksheet.write(row, 5, '', formats['input_text'])
                    
                    row += 1
    
    def _create_validation_sheet(self, workbook, formats: Dict):
        """Create sheet with validation rules and data types."""
        worksheet = workbook.add_worksheet('Validation Rules')
        
        # Set column widths
        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 40)
        worksheet.set_column('C:C', 30)
        
        # Title
        worksheet.merge_range('A1:C1', 'Data Validation Rules', formats['title'])
        
        # Validation rules
        rules = [
            ('Data Type', 'Description', 'Example'),
            ('numeric', 'Positive numbers only', '1234.56'),
            ('integer', 'Whole numbers only', '100'),
            ('percentage', 'Values between 0-100', '75.5'),
            ('currency', 'Monetary values', '$10,000.00'),
            ('boolean', 'Yes/No values', 'Yes'),
            ('text', 'Any text input', 'Description text'),
        ]
        
        row = 2
        for rule in rules:
            if row == 2:  # Header row
                for col, value in enumerate(rule):
                    worksheet.write(row, col, value, formats['header'])
            else:
                for col, value in enumerate(rule):
                    worksheet.write(row, col, value, formats['metric'])
            row += 1
        
        # Additional notes
        row += 2
        worksheet.merge_range(f'A{row+1}:C{row+1}', 'Notes', formats['header'])
        row += 2
        
        notes = [
            "• Yellow cells require data input",
            "• Ensure units match exactly as specified",
            "• Use consistent reporting periods across all metrics",
            "• Leave cells blank if data is not available",
            "• Add explanatory notes where necessary"
        ]
        
        for note in notes:
            worksheet.write(row, 0, note, formats['note'])
            row += 1
    
    def _get_input_format(self, formats: Dict, data_type: str):
        """Get the appropriate format for input cells based on data type."""
        format_map = {
            'numeric': 'input_numeric',
            'integer': 'input_numeric',
            'percentage': 'input_percent',
            'currency': 'input_currency',
            'text': 'input_text',
            'boolean': 'input_text'
        }
        return formats.get(format_map.get(data_type, 'input_text'))
    
    def _add_validation(self, worksheet, row: int, col: int, data_type: str):
        """Add data validation to cells based on data type."""
        if data_type == 'numeric' or data_type == 'integer':
            worksheet.data_validation(row, col, row, col, {
                'validate': 'decimal',
                'criteria': '>=',
                'value': 0,
                'error_message': 'Please enter a positive number'
            })
        elif data_type == 'percentage':
            worksheet.data_validation(row, col, row, col, {
                'validate': 'decimal',
                'criteria': 'between',
                'minimum': 0,
                'maximum': 100,
                'error_message': 'Please enter a value between 0 and 100'
            })
        elif data_type == 'boolean':
            worksheet.data_validation(row, col, row, col, {
                'validate': 'list',
                'source': ['Yes', 'No', 'N/A'],
                'dropdown': True
            })