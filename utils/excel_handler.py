import pandas as pd
from datetime import datetime
import json
import re

class ExcelGenerator:
    def create_form(self, crew_result, investor_name, property_list):
        """Generate Excel form from crew analysis results"""
        
        # Parse the crew results (this is simplified - you'd parse the actual structure)
        metrics = self._parse_metrics(crew_result)
        
        # Create Excel file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"output/ESG_Form_{investor_name}_{timestamp}.xlsx"
        
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Formatting
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4CAF50',
                'font_color': 'white',
                'border': 1
            })
            
            # Create overview sheet
            overview_data = {
                'Investor': [investor_name],
                'Properties': [', '.join(property_list)],
                'Generated': [datetime.now().strftime("%Y-%m-%d")],
                'Total Metrics': [len(metrics)]
            }
            
            overview_df = pd.DataFrame(overview_data)
            overview_df.to_excel(writer, sheet_name='Overview', index=False)
            
            # Create sheet for each property
            for property in property_list:
                self._create_property_sheet(
                    writer, workbook, property, metrics, header_format
                )
        
        return filename
    
    def _create_property_sheet(self, writer, workbook, property_name, metrics, header_format):
        """Create a sheet for each property"""
        
        # Structure: Category | Metric | Unit | Value | Notes
        data = []
        
        for category, category_metrics in metrics.items():
            for metric in category_metrics:
                data.append({
                    'Category': category,
                    'Metric': metric['name'],
                    'Unit': metric.get('unit', ''),
                    'Value': '',  # Empty for data entry
                    'Notes': metric.get('description', '')
                })
        
        df = pd.DataFrame(data)
        df.to_excel(writer, sheet_name=property_name[:31], index=False)  # Excel limit
        
        # Add data validation
        worksheet = writer.sheets[property_name[:31]]
        
        # Format headers
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Add dropdowns for certain metrics
        for row_num, row in df.iterrows():
            if 'yes/no' in str(row['Unit']).lower():
                worksheet.data_validation(
                    row_num + 1, 3, row_num + 1, 3,
                    {'validate': 'list', 'source': ['Yes', 'No', 'N/A']}
                )
    
    def _parse_metrics(self, crew_result):
        """Parse metrics from crew analysis results"""
        # This is a simplified parser - you'd implement based on your crew output format
        
        # Example structure
        metrics = {
            'Energy & Emissions': [
                {'name': 'Total Energy Consumption', 'unit': 'MWh/year'},
                {'name': 'Renewable Energy %', 'unit': 'Percentage'},
                {'name': 'Scope 1 Emissions', 'unit': 'tCO2e'},
                {'name': 'Scope 2 Emissions', 'unit': 'tCO2e'},
            ],
            'Water Management': [
                {'name': 'Water Consumption', 'unit': 'm³/year'},
                {'name': 'Water Recycling Rate', 'unit': 'Percentage'},
            ],
            'Waste Management': [
                {'name': 'Total Waste Generated', 'unit': 'tonnes/year'},
                {'name': 'Recycling Rate', 'unit': 'Percentage'},
            ],
            'Social Impact': [
                {'name': 'Community Engagement Programs', 'unit': 'Yes/No'},
                {'name': 'Local Employment Rate', 'unit': 'Percentage'},
            ]
        }
        
        return metrics