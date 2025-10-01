# Quick Start Guide

## Get Started in 3 Steps

### Step 1: Install Dependencies

```bash
pip install -r requirements.txt
```

### Step 2: Generate Dashboards

```bash
python dashboard_generator.py
```

This creates three Excel files in the `output/` directory:
- `issue_tracker_dashboard.xlsx`
- `client_launch_tracker.xlsx`
- `integrated_dashboard.xlsx`

### Step 3: Open and Use

Open any of the generated Excel files in Microsoft Excel, LibreOffice Calc, or Google Sheets.

## What's Inside

### Issue Tracker Dashboard
- **Issue Tracker** sheet: Track all project issues
- **Issue Summary** sheet: Statistics and charts
- **KPI Dashboard** sheet: Key performance metrics

### Client Launch Tracker
- **Launch Tracker** sheet: Monitor client projects
- **Timeline** sheet: Track milestones and deadlines

### Integrated Dashboard
- **Dashboard Overview**: Quick stats and navigation
- **Issue Tracker**: Full issue management
- **Launch Tracker**: Complete project tracking
- **KPI Dashboard**: Comprehensive metrics

## Customization

### Add Your Own Data

1. Open the Excel file
2. Navigate to the appropriate sheet
3. Add new rows below the sample data
4. The formatting will automatically apply

### Change Colors

Edit `config.py` to customize color schemes:
```python
COLORS = {
    'primary': '366092',    # Blue
    'danger': 'FF0000',     # Red
    'success': '00B050',    # Green
    # ... more colors
}
```

### Regenerate Dashboards

After modifying the Python scripts:
```bash
python dashboard_generator.py
```

## Tips

1. **Use Filters**: Click the filter buttons in Excel to sort and filter data
2. **Freeze Panes**: The header row is frozen for easy scrolling
3. **Color Coding**: Pay attention to color-coded priorities and statuses
4. **Charts**: Summary charts update automatically when you change data
5. **Templates**: Keep a backup of the original files as templates

## Need Help?

Check the main [README.md](README.md) for detailed documentation.
