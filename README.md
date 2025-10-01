# FASTProject

Excel Dashboard System for Issue Tracking, Client Launch Management, and KPI Monitoring

## Overview

FASTProject is a comprehensive dashboard system that provides integrated Excel-based tools for:
- **Issue Tracking Dashboard**: Track and manage project issues, bugs, and tasks
- **Client Launch Tracker**: Monitor client project launches and milestones
- **KPI Dashboard**: View key performance indicators and project metrics

## Features

### Issue Tracking Dashboard
- Complete issue lifecycle management
- Priority-based color coding (Critical, High, Medium, Low)
- Status tracking (Open, In Progress, Resolved, Closed)
- Assignment and reporting tracking
- Category and tag management
- Visual summary with charts

### Client Launch Tracker
- Client project timeline management
- Budget tracking
- Completion percentage monitoring
- Phase tracking
- Risk level assessment with color coding
- Milestone tracking

### KPI Dashboard
- Real-time project statistics
- Weekly trend analysis
- Visual charts and graphs
- Performance metrics
- Resolution time tracking
- On-time delivery rate monitoring

## Installation

1. Clone the repository:
```bash
git clone https://github.com/ninirob/FASTProject.git
cd FASTProject
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Generate Dashboards

Run the dashboard generator to create Excel files:

```bash
python dashboard_generator.py
```

This will create three Excel files in the `output/` directory:
1. `issue_tracker_dashboard.xlsx` - Issue tracking with KPI dashboard
2. `client_launch_tracker.xlsx` - Client launch management
3. `integrated_dashboard.xlsx` - Combined dashboard with all components

### Test Sample Data

Generate sample data for testing:

```bash
python sample_data.py
```

## File Structure

```
FASTProject/
├── dashboard_generator.py    # Main dashboard generator
├── sample_data.py            # Sample data generator
├── config.py                 # Configuration settings
├── requirements.txt          # Python dependencies
├── README.md                 # This file
└── output/                   # Generated Excel files (created automatically)
    ├── issue_tracker_dashboard.xlsx
    ├── client_launch_tracker.xlsx
    └── integrated_dashboard.xlsx
```

## Dashboard Components

### 1. Issue Tracker Dashboard

**Sheets:**
- **Issue Tracker**: Main tracking sheet with all issues
  - Columns: Issue ID, Title, Description, Priority, Status, Assigned To, Reporter, Dates, Category, Tags
  - Color-coded priorities and statuses
  - Sortable and filterable

- **Issue Summary**: Statistical overview
  - Issues by Status (with pie chart)
  - Issues by Priority
  - Quick statistics

- **KPI Dashboard**: Key performance indicators
  - Total/Open/Resolved issue counts
  - Average resolution time
  - Weekly trends with bar chart
  - On-time delivery rate

### 2. Client Launch Tracker

**Sheets:**
- **Launch Tracker**: Main project tracking
  - Columns: Client Name, Project Name, Launch Date, Status, PM, Budget, Completion %, Phase, Risk Level, Notes
  - Status color coding (Planning, On Track, At Risk, Completed)
  - Risk level indicators

- **Timeline**: Milestone tracking
  - Client milestones
  - Target vs. Actual dates
  - Progress monitoring

### 3. Integrated Dashboard

**Sheets:**
- **Dashboard Overview**: Central hub with quick stats and navigation
- **Issue Tracker**: Full issue tracking functionality
- **Launch Tracker**: Complete client launch management
- **KPI Dashboard**: Comprehensive metrics and analytics

## Customization

### Colors and Styles

Edit `config.py` to customize:
- Color schemes
- Priority levels
- Status definitions
- Risk levels
- Column configurations

### Data Structure

Modify `dashboard_generator.py` to:
- Add/remove columns
- Change sample data
- Adjust chart types
- Customize formatting

## Requirements

- Python 3.7+
- openpyxl 3.1.2+
- pandas 2.0.0+
- numpy 1.24.0+

## Best Practices

1. **Regular Updates**: Update dashboards regularly to maintain accuracy
2. **Backup Data**: Keep backups of Excel files before major changes
3. **Use Filters**: Utilize Excel's filter features to analyze specific data
4. **Export Reports**: Generate reports for stakeholders as needed
5. **Consistent Data Entry**: Follow consistent naming conventions

## Integration

The dashboards are designed to work together:
- Issues can reference client projects
- KPIs aggregate data from both tracking systems
- Integrated dashboard provides unified view
- Data can be linked across sheets using Excel formulas

## Troubleshooting

### Common Issues

1. **Import Errors**: Ensure all dependencies are installed
   ```bash
   pip install -r requirements.txt
   ```

2. **Permission Errors**: Ensure write permissions for the output directory

3. **Excel Opens Read-Only**: Close any open Excel files before regenerating

## Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for bugs and feature requests.

## License

This project is open source and available for use and modification.

## Support

For questions or issues, please open an issue on GitHub.