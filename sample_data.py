"""
Sample data generator for testing dashboards
"""

from datetime import datetime, timedelta
import random


class SampleDataGenerator:
    """Generate sample data for testing dashboards"""
    
    def __init__(self):
        self.priorities = ['Critical', 'High', 'Medium', 'Low']
        self.issue_statuses = ['Open', 'In Progress', 'Resolved', 'Closed']
        self.launch_statuses = ['Planning', 'On Track', 'At Risk', 'Completed']
        self.risk_levels = ['Low', 'Medium', 'High']
        self.categories = ['Technical', 'UI/UX', 'Performance', 'Security', 'Documentation', 'Business']
        self.phases = ['Planning', 'Development', 'Testing', 'Deployment', 'Maintenance', 'Migration', 'Completed']
        
        self.team_members = [
            'John Doe', 'Jane Smith', 'Alice Johnson', 'Bob Wilson',
            'Charlie Brown', 'David Lee', 'Eve Davis', 'Frank Martin',
            'Grace Chen', 'Henry Taylor'
        ]
        
        self.clients = [
            'Acme Corp', 'TechStart Inc', 'Global Solutions', 'InnovateCo',
            'DataCorp', 'CloudSystems', 'NextGen Ltd', 'FutureTech',
            'Enterprise Solutions', 'Digital Dynamics'
        ]
        
        self.project_types = [
            'CRM Implementation', 'Mobile App Launch', 'Cloud Migration',
            'Website Redesign', 'Analytics Platform', 'API Integration',
            'Database Upgrade', 'Security Audit', 'System Modernization'
        ]
    
    def generate_issues(self, count=10):
        """Generate sample issue data"""
        issues = []
        base_date = datetime.now() - timedelta(days=30)
        
        for i in range(count):
            issue_id = f'ISS-{str(i+1).zfill(3)}'
            priority = random.choice(self.priorities)
            status = random.choice(self.issue_statuses)
            category = random.choice(self.categories)
            
            created_date = base_date + timedelta(days=random.randint(0, 30))
            due_date = created_date + timedelta(days=random.randint(3, 14))
            
            resolution_date = ''
            if status in ['Resolved', 'Closed']:
                resolution_date = created_date + timedelta(days=random.randint(1, 10))
            
            issue = {
                'issue_id': issue_id,
                'title': f'{category} Issue {i+1}',
                'description': f'Description for {category.lower()} issue',
                'priority': priority,
                'status': status,
                'assigned_to': random.choice(self.team_members),
                'reporter': random.choice(self.team_members),
                'created_date': created_date.strftime('%Y-%m-%d'),
                'due_date': due_date.strftime('%Y-%m-%d'),
                'resolution_date': resolution_date.strftime('%Y-%m-%d') if resolution_date else '',
                'category': category,
                'tags': f'{category.lower()},priority-{priority.lower()}'
            }
            issues.append(issue)
        
        return issues
    
    def generate_client_launches(self, count=8):
        """Generate sample client launch data"""
        launches = []
        base_date = datetime.now()
        
        for i in range(count):
            client = random.choice(self.clients)
            project = random.choice(self.project_types)
            status = random.choice(self.launch_statuses)
            risk = random.choice(self.risk_levels)
            phase = random.choice(self.phases)
            
            launch_date = base_date + timedelta(days=random.randint(-30, 90))
            
            completion_pct = 0
            if status == 'Completed':
                completion_pct = 100
            elif status == 'On Track':
                completion_pct = random.randint(40, 85)
            elif status == 'At Risk':
                completion_pct = random.randint(20, 60)
            else:  # Planning
                completion_pct = random.randint(0, 30)
            
            budget = random.randint(50, 300) * 1000
            
            launch = {
                'client_name': client,
                'project_name': project,
                'launch_date': launch_date.strftime('%Y-%m-%d'),
                'status': status,
                'project_manager': random.choice(self.team_members),
                'budget': f'${budget:,}',
                'completion_pct': completion_pct,
                'phase': phase,
                'risk_level': risk,
                'notes': f'Project progressing in {phase.lower()} phase'
            }
            launches.append(launch)
        
        return launches
    
    def generate_kpi_data(self):
        """Generate sample KPI data"""
        total_issues = random.randint(20, 50)
        open_issues = random.randint(5, 15)
        resolved_issues = random.randint(10, 25)
        critical_issues = random.randint(1, 5)
        
        avg_resolution_time = round(random.uniform(2.0, 5.0), 1)
        on_time_rate = random.randint(75, 95)
        
        kpis = {
            'total_issues': total_issues,
            'open_issues': open_issues,
            'resolved_issues': resolved_issues,
            'critical_issues': critical_issues,
            'avg_resolution_time': f'{avg_resolution_time} days',
            'on_time_rate': f'{on_time_rate}%'
        }
        
        # Weekly trends
        weeks = []
        for i in range(4):
            week_data = {
                'week': f'Week {i+1}',
                'created': random.randint(3, 8),
                'resolved': random.randint(2, 6),
                'open': random.randint(2, 5)
            }
            weeks.append(week_data)
        
        kpis['weekly_trends'] = weeks
        
        return kpis


def main():
    """Test sample data generation"""
    generator = SampleDataGenerator()
    
    print("Sample Issues:")
    issues = generator.generate_issues(5)
    for issue in issues:
        print(f"  {issue['issue_id']}: {issue['title']} - {issue['status']}")
    
    print("\nSample Client Launches:")
    launches = generator.generate_client_launches(5)
    for launch in launches:
        print(f"  {launch['client_name']}: {launch['project_name']} - {launch['status']}")
    
    print("\nSample KPIs:")
    kpis = generator.generate_kpi_data()
    print(f"  Total Issues: {kpis['total_issues']}")
    print(f"  Open Issues: {kpis['open_issues']}")
    print(f"  Average Resolution Time: {kpis['avg_resolution_time']}")


if __name__ == '__main__':
    main()
