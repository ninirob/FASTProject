"""
Configuration settings for FAST Project Dashboards
"""

# Dashboard Color Schemes
COLORS = {
    'primary': '366092',
    'secondary': '2E75B6',
    'success': '00B050',
    'warning': 'FFC000',
    'danger': 'FF0000',
    'info': '87CEEB',
}

# Priority Levels
PRIORITIES = {
    'Critical': {'color': 'FF0000', 'text_color': 'FFFFFF'},
    'High': {'color': 'FFA500', 'text_color': '000000'},
    'Medium': {'color': 'FFFF00', 'text_color': '000000'},
    'Low': {'color': '90EE90', 'text_color': '000000'},
}

# Status Definitions
ISSUE_STATUSES = {
    'Open': {'color': 'FFE4B5', 'text_color': '000000'},
    'In Progress': {'color': '87CEEB', 'text_color': '000000'},
    'Resolved': {'color': '00FF00', 'text_color': '000000'},
    'Closed': {'color': '808080', 'text_color': 'FFFFFF'},
}

LAUNCH_STATUSES = {
    'Planning': {'color': 'FFC000', 'text_color': '000000'},
    'On Track': {'color': '92D050', 'text_color': '000000'},
    'At Risk': {'color': 'FF0000', 'text_color': 'FFFFFF'},
    'Completed': {'color': '00B050', 'text_color': 'FFFFFF'},
}

# Risk Levels
RISK_LEVELS = {
    'Low': {'color': '92D050', 'text_color': '000000'},
    'Medium': {'color': 'FFC000', 'text_color': '000000'},
    'High': {'color': 'FF0000', 'text_color': 'FFFFFF'},
}

# Output Settings
OUTPUT_DIR = 'output'
DEFAULT_FILENAMES = {
    'issue_tracker': 'issue_tracker_dashboard.xlsx',
    'client_launch': 'client_launch_tracker.xlsx',
    'integrated': 'integrated_dashboard.xlsx',
}

# Column Configurations
ISSUE_TRACKER_COLUMNS = [
    'Issue ID', 'Title', 'Description', 'Priority', 'Status',
    'Assigned To', 'Reporter', 'Created Date', 'Due Date',
    'Resolution Date', 'Category', 'Tags'
]

LAUNCH_TRACKER_COLUMNS = [
    'Client Name', 'Project Name', 'Launch Date', 'Status',
    'Project Manager', 'Budget', 'Completion %', 'Phase',
    'Risk Level', 'Notes'
]
