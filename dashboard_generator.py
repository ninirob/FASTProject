"""
Excel Dashboard Generator for Issue Tracking, Client Launch Tracker, and KPI Dashboard
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import os


class DashboardGenerator:
    """Generate Excel dashboards for project management"""
    
    def __init__(self, output_dir='output'):
        self.output_dir = output_dir
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
    
    def create_issue_tracker(self, filename='issue_tracker_dashboard.xlsx'):
        """Create Issue Tracking Dashboard"""
        wb = openpyxl.Workbook()
        
        # Remove default sheet
        wb.remove(wb.active)
        
        # Create Issue Tracker sheet
        ws_issues = wb.create_sheet('Issue Tracker')
        self._setup_issue_tracker_sheet(ws_issues)
        
        # Create Issue Summary sheet
        ws_summary = wb.create_sheet('Issue Summary')
        self._setup_issue_summary_sheet(ws_summary)
        
        # Create KPI Dashboard sheet
        ws_kpi = wb.create_sheet('KPI Dashboard')
        self._setup_kpi_dashboard_sheet(ws_kpi)
        
        # Save workbook
        filepath = os.path.join(self.output_dir, filename)
        wb.save(filepath)
        print(f"Issue Tracker Dashboard created: {filepath}")
        return filepath
    
    def create_client_launch_tracker(self, filename='client_launch_tracker.xlsx'):
        """Create Client Launch Tracker"""
        wb = openpyxl.Workbook()
        
        # Remove default sheet
        wb.remove(wb.active)
        
        # Create Launch Tracker sheet
        ws_launch = wb.create_sheet('Launch Tracker')
        self._setup_launch_tracker_sheet(ws_launch)
        
        # Create Timeline sheet
        ws_timeline = wb.create_sheet('Timeline')
        self._setup_timeline_sheet(ws_timeline)
        
        # Save workbook
        filepath = os.path.join(self.output_dir, filename)
        wb.save(filepath)
        print(f"Client Launch Tracker created: {filepath}")
        return filepath
    
    def create_integrated_dashboard(self, filename='integrated_dashboard.xlsx'):
        """Create Integrated Dashboard with all components"""
        wb = openpyxl.Workbook()
        
        # Remove default sheet
        wb.remove(wb.active)
        
        # Create Dashboard Overview
        ws_overview = wb.create_sheet('Dashboard Overview')
        self._setup_dashboard_overview(ws_overview)
        
        # Create Issue Tracker sheet
        ws_issues = wb.create_sheet('Issue Tracker')
        self._setup_issue_tracker_sheet(ws_issues)
        
        # Create Launch Tracker sheet
        ws_launch = wb.create_sheet('Launch Tracker')
        self._setup_launch_tracker_sheet(ws_launch)
        
        # Create KPI Dashboard sheet
        ws_kpi = wb.create_sheet('KPI Dashboard')
        self._setup_kpi_dashboard_sheet(ws_kpi)
        
        # Save workbook
        filepath = os.path.join(self.output_dir, filename)
        wb.save(filepath)
        print(f"Integrated Dashboard created: {filepath}")
        return filepath
    
    def _setup_issue_tracker_sheet(self, ws):
        """Setup Issue Tracker sheet structure"""
        # Define headers
        headers = [
            'Issue ID', 'Title', 'Description', 'Priority', 'Status',
            'Assigned To', 'Reporter', 'Created Date', 'Due Date',
            'Resolution Date', 'Category', 'Tags'
        ]
        
        # Apply header styling
        header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=11)
        header_alignment = Alignment(horizontal='center', vertical='center')
        
        # Write headers
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
        
        # Set column widths
        column_widths = [10, 30, 40, 12, 12, 15, 15, 15, 15, 15, 15, 20]
        for col, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = width
        
        # Add sample data
        sample_data = [
            ['ISS-001', 'Database Connection Issue', 'Unable to connect to production database', 'High', 'Open',
             'John Doe', 'Jane Smith', '2024-01-15', '2024-01-20', '', 'Technical', 'database,urgent'],
            ['ISS-002', 'UI Button Alignment', 'Submit button misaligned on mobile', 'Medium', 'In Progress',
             'Alice Johnson', 'Bob Wilson', '2024-01-16', '2024-01-25', '', 'UI/UX', 'frontend,mobile'],
            ['ISS-003', 'Performance Optimization', 'Page load time exceeds 3 seconds', 'High', 'Open',
             'Charlie Brown', 'David Lee', '2024-01-17', '2024-01-30', '', 'Performance', 'optimization'],
            ['ISS-004', 'Documentation Update', 'API documentation needs updating', 'Low', 'Resolved',
             'Eve Davis', 'Frank Martin', '2024-01-10', '2024-01-18', '2024-01-18', 'Documentation', 'docs'],
            ['ISS-005', 'Security Vulnerability', 'SQL injection risk in search feature', 'Critical', 'In Progress',
             'John Doe', 'Security Team', '2024-01-18', '2024-01-22', '', 'Security', 'security,critical'],
        ]
        
        for row_idx, row_data in enumerate(sample_data, 2):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.alignment = Alignment(vertical='center')
                
                # Apply conditional formatting for priority
                if col_idx == 4:  # Priority column
                    if value == 'Critical':
                        cell.fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
                        cell.font = Font(color='FFFFFF', bold=True)
                    elif value == 'High':
                        cell.fill = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
                    elif value == 'Medium':
                        cell.fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
                    elif value == 'Low':
                        cell.fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
                
                # Apply conditional formatting for status
                if col_idx == 5:  # Status column
                    if value == 'Resolved':
                        cell.fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
                    elif value == 'In Progress':
                        cell.fill = PatternFill(start_color='87CEEB', end_color='87CEEB', fill_type='solid')
                    elif value == 'Open':
                        cell.fill = PatternFill(start_color='FFE4B5', end_color='FFE4B5', fill_type='solid')
        
        # Freeze panes
        ws.freeze_panes = 'A2'
    
    def _setup_issue_summary_sheet(self, ws):
        """Setup Issue Summary sheet with statistics"""
        title_font = Font(bold=True, size=14)
        header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF')
        
        # Title
        ws['A1'] = 'Issue Tracker Summary'
        ws['A1'].font = title_font
        ws.merge_cells('A1:D1')
        
        # Summary by Status
        ws['A3'] = 'Status'
        ws['B3'] = 'Count'
        ws['A3'].fill = header_fill
        ws['B3'].fill = header_fill
        ws['A3'].font = header_font
        ws['B3'].font = header_font
        
        status_data = [
            ['Open', 2],
            ['In Progress', 2],
            ['Resolved', 1],
        ]
        
        for idx, (status, count) in enumerate(status_data, 4):
            ws[f'A{idx}'] = status
            ws[f'B{idx}'] = count
        
        # Summary by Priority
        ws['A8'] = 'Priority'
        ws['B8'] = 'Count'
        ws['A8'].fill = header_fill
        ws['B8'].fill = header_fill
        ws['A8'].font = header_font
        ws['B8'].font = header_font
        
        priority_data = [
            ['Critical', 1],
            ['High', 2],
            ['Medium', 1],
            ['Low', 1],
        ]
        
        for idx, (priority, count) in enumerate(priority_data, 9):
            ws[f'A{idx}'] = priority
            ws[f'B{idx}'] = count
        
        # Set column widths
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 10
        
        # Add pie chart for status distribution
        pie = PieChart()
        pie.title = "Issues by Status"
        labels = Reference(ws, min_col=1, min_row=4, max_row=6)
        data = Reference(ws, min_col=2, min_row=3, max_row=6)
        pie.add_data(data, titles_from_data=True)
        pie.set_categories(labels)
        pie.height = 10
        pie.width = 15
        ws.add_chart(pie, "D3")
    
    def _setup_kpi_dashboard_sheet(self, ws):
        """Setup KPI Dashboard sheet"""
        title_font = Font(bold=True, size=16)
        header_fill = PatternFill(start_color='366092', end_color='366092', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=12)
        kpi_font = Font(bold=True, size=20)
        
        # Title
        ws['A1'] = 'KPI Dashboard'
        ws['A1'].font = title_font
        ws.merge_cells('A1:F1')
        
        # KPI Metrics
        kpis = [
            ('Total Issues', 5, 'B3'),
            ('Open Issues', 2, 'D3'),
            ('Resolved Issues', 1, 'F3'),
            ('Average Resolution Time', '3.2 days', 'B6'),
            ('Critical Issues', 1, 'D6'),
            ('On-Time Delivery Rate', '85%', 'F6'),
        ]
        
        row = 3
        for label, value, cell_ref in kpis:
            # Label
            ws[cell_ref] = label
            ws[cell_ref].font = header_font
            ws[cell_ref].fill = header_fill
            ws[cell_ref].alignment = Alignment(horizontal='center')
            
            # Value (next row)
            value_cell = ws[cell_ref[0] + str(int(cell_ref[1]) + 1)]
            value_cell.value = value
            value_cell.font = kpi_font
            value_cell.alignment = Alignment(horizontal='center')
        
        # Set column widths
        for col in ['A', 'B', 'C', 'D', 'E', 'F']:
            ws.column_dimensions[col].width = 18
        
        # Trends section
        ws['A9'] = 'Weekly Trends'
        ws['A9'].font = Font(bold=True, size=14)
        ws.merge_cells('A9:F9')
        
        trends_headers = ['Week', 'Issues Created', 'Issues Resolved', 'Issues Open']
        for col, header in enumerate(trends_headers, 1):
            cell = ws.cell(row=10, column=col, value=header)
            cell.fill = header_fill
            cell.font = header_font
        
        # Sample trend data
        trends_data = [
            ['Week 1', 3, 1, 2],
            ['Week 2', 2, 2, 2],
            ['Week 3', 4, 3, 3],
            ['Week 4', 5, 4, 4],
        ]
        
        for row_idx, row_data in enumerate(trends_data, 11):
            for col_idx, value in enumerate(row_data, 1):
                ws.cell(row=row_idx, column=col_idx, value=value)
        
        # Add bar chart for trends
        chart = BarChart()
        chart.title = "Weekly Issue Trends"
        chart.y_axis.title = 'Count'
        chart.x_axis.title = 'Week'
        
        data = Reference(ws, min_col=2, min_row=10, max_row=14, max_col=4)
        cats = Reference(ws, min_col=1, min_row=11, max_row=14)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.height = 12
        chart.width = 20
        ws.add_chart(chart, "A16")
    
    def _setup_launch_tracker_sheet(self, ws):
        """Setup Client Launch Tracker sheet"""
        # Define headers
        headers = [
            'Client Name', 'Project Name', 'Launch Date', 'Status',
            'Project Manager', 'Budget', 'Completion %', 'Phase',
            'Risk Level', 'Notes'
        ]
        
        # Apply header styling
        header_fill = PatternFill(start_color='2E75B6', end_color='2E75B6', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=11)
        header_alignment = Alignment(horizontal='center', vertical='center')
        
        # Write headers
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
        
        # Set column widths
        column_widths = [20, 25, 15, 15, 20, 15, 15, 15, 12, 30]
        for col, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = width
        
        # Add sample data
        sample_data = [
            ['Acme Corp', 'CRM Implementation', '2024-03-15', 'On Track',
             'Sarah Johnson', '$150,000', 75, 'Development', 'Low', 'Progressing well'],
            ['TechStart Inc', 'Mobile App Launch', '2024-02-28', 'At Risk',
             'Mike Chen', '$80,000', 45, 'Testing', 'High', 'Resource constraints'],
            ['Global Solutions', 'Cloud Migration', '2024-04-30', 'On Track',
             'Emily Rodriguez', '$250,000', 60, 'Migration', 'Medium', 'On schedule'],
            ['InnovateCo', 'Website Redesign', '2024-02-15', 'Completed',
             'Tom Anderson', '$50,000', 100, 'Completed', 'Low', 'Successfully launched'],
            ['DataCorp', 'Analytics Platform', '2024-05-15', 'Planning',
             'Lisa Wang', '$200,000', 20, 'Planning', 'Medium', 'Requirements gathering'],
        ]
        
        for row_idx, row_data in enumerate(sample_data, 2):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.alignment = Alignment(vertical='center')
                
                # Apply conditional formatting for status
                if col_idx == 4:  # Status column
                    if value == 'Completed':
                        cell.fill = PatternFill(start_color='00B050', end_color='00B050', fill_type='solid')
                        cell.font = Font(color='FFFFFF', bold=True)
                    elif value == 'On Track':
                        cell.fill = PatternFill(start_color='92D050', end_color='92D050', fill_type='solid')
                    elif value == 'At Risk':
                        cell.fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
                        cell.font = Font(color='FFFFFF', bold=True)
                    elif value == 'Planning':
                        cell.fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
                
                # Apply conditional formatting for risk level
                if col_idx == 9:  # Risk Level column
                    if value == 'High':
                        cell.fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
                        cell.font = Font(color='FFFFFF', bold=True)
                    elif value == 'Medium':
                        cell.fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
                    elif value == 'Low':
                        cell.fill = PatternFill(start_color='92D050', end_color='92D050', fill_type='solid')
        
        # Freeze panes
        ws.freeze_panes = 'A2'
    
    def _setup_timeline_sheet(self, ws):
        """Setup Timeline sheet for launch tracking"""
        title_font = Font(bold=True, size=14)
        
        # Title
        ws['A1'] = 'Client Launch Timeline'
        ws['A1'].font = title_font
        ws.merge_cells('A1:D1')
        
        # Timeline data
        headers = ['Client', 'Milestone', 'Target Date', 'Actual Date']
        header_fill = PatternFill(start_color='2E75B6', end_color='2E75B6', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF')
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col, value=header)
            cell.fill = header_fill
            cell.font = header_font
        
        timeline_data = [
            ['Acme Corp', 'Kickoff', '2024-01-05', '2024-01-05'],
            ['Acme Corp', 'Design Complete', '2024-01-20', '2024-01-22'],
            ['Acme Corp', 'Development Complete', '2024-02-28', ''],
            ['Acme Corp', 'Launch', '2024-03-15', ''],
            ['TechStart Inc', 'Kickoff', '2023-12-01', '2023-12-01'],
            ['TechStart Inc', 'Beta Release', '2024-01-30', '2024-02-05'],
            ['TechStart Inc', 'Launch', '2024-02-28', ''],
        ]
        
        for row_idx, row_data in enumerate(timeline_data, 4):
            for col_idx, value in enumerate(row_data, 1):
                ws.cell(row=row_idx, column=col_idx, value=value)
        
        # Set column widths
        for col in ['A', 'B', 'C', 'D']:
            ws.column_dimensions[col].width = 20
    
    def _setup_dashboard_overview(self, ws):
        """Setup Dashboard Overview sheet"""
        title_font = Font(bold=True, size=18, color='366092')
        section_font = Font(bold=True, size=14)
        
        # Main title
        ws['A1'] = 'FAST Project Management Dashboard'
        ws['A1'].font = title_font
        ws.merge_cells('A1:F1')
        ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 30
        
        # Description
        ws['A3'] = 'Integrated Dashboard for Issue Tracking, Client Launch Management, and KPI Monitoring'
        ws.merge_cells('A3:F3')
        ws['A3'].alignment = Alignment(horizontal='center')
        
        # Dashboard Sections
        ws['A5'] = 'Dashboard Components:'
        ws['A5'].font = section_font
        
        components = [
            ('Issue Tracker', 'Track and manage project issues, bugs, and tasks'),
            ('Launch Tracker', 'Monitor client project launches and milestones'),
            ('KPI Dashboard', 'View key performance indicators and metrics'),
        ]
        
        row = 6
        for component, description in components:
            ws[f'A{row}'] = f'• {component}'
            ws[f'A{row}'].font = Font(bold=True, size=11)
            ws[f'B{row}'] = description
            ws.merge_cells(f'B{row}:F{row}')
            row += 1
        
        # Quick Stats
        ws['A10'] = 'Quick Statistics:'
        ws['A10'].font = section_font
        
        stats_fill = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
        stats = [
            ('Total Active Issues:', '4'),
            ('Active Client Projects:', '4'),
            ('Projects On Track:', '3'),
            ('Critical Issues:', '1'),
        ]
        
        row = 11
        for label, value in stats:
            ws[f'A{row}'] = label
            ws[f'B{row}'] = value
            ws[f'A{row}'].fill = stats_fill
            ws[f'B{row}'].fill = stats_fill
            ws[f'B{row}'].font = Font(bold=True, size=12)
            row += 1
        
        # Set column widths
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 40
        
        # Instructions
        ws['A17'] = 'Instructions:'
        ws['A17'].font = section_font
        
        instructions = [
            'Navigate between sheets using the tabs at the bottom',
            'Update data regularly to maintain accuracy',
            'Use filters to analyze specific data points',
            'Export reports as needed for stakeholders',
        ]
        
        row = 18
        for instruction in instructions:
            ws[f'A{row}'] = f'• {instruction}'
            ws.merge_cells(f'A{row}:F{row}')
            row += 1


def main():
    """Main function to generate all dashboards"""
    print("=" * 60)
    print("FAST Project Dashboard Generator")
    print("=" * 60)
    
    generator = DashboardGenerator()
    
    # Generate individual dashboards
    print("\nGenerating dashboards...")
    generator.create_issue_tracker()
    generator.create_client_launch_tracker()
    
    # Generate integrated dashboard
    generator.create_integrated_dashboard()
    
    print("\n" + "=" * 60)
    print("Dashboard generation completed successfully!")
    print("=" * 60)
    print("\nGenerated files:")
    print("1. output/issue_tracker_dashboard.xlsx - Issue tracking with KPI dashboard")
    print("2. output/client_launch_tracker.xlsx - Client launch management")
    print("3. output/integrated_dashboard.xlsx - Combined dashboard with all components")
    print("\nYou can now open these Excel files and customize them as needed.")


if __name__ == '__main__':
    main()
