import pandas as pd

class ReportManager:
    def __init__(self, meeting_manager, report_text_edit):
        self.meeting_manager = meeting_manager
        self.report_text_edit = report_text_edit

    def update_report(self):
        self.meeting_manager.check_unassigned()
        
        df_filtered = self.meeting_manager.df[~self.meeting_manager.df['Subject'].isin(self.meeting_manager.removed_meetings)]
        df_filtered['Day'] = df_filtered['Start'].dt.day_name()
        day_order = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
        df_filtered['Day'] = pd.Categorical(df_filtered['Day'], categories=day_order, ordered=True)

        time_by_project_day = df_filtered.groupby(['Project', 'Day'], observed=False)['Duration'].sum().unstack().fillna(0)
        time_by_project_day = time_by_project_day[day_order].fillna(0)

        # Create the report in HTML format for better alignment
        report_html = """
        <table border="1" cellpadding="3" cellspacing="0">
            <tr>
                <th>Project</th>
                <th>Sunday</th>
                <th>Monday</th>
                <th>Tuesday</th>
                <th>Wednesday</th>
                <th>Thursday</th>
                <th>Friday</th>
                <th>Saturday</th>
            </tr>
        """

        for project, row in time_by_project_day.iterrows():
            row_style = 'style="background-color: #ffcccc;"' if project == "Unassigned" else ""
            report_html += f"<tr {row_style}><td>{project}</td>"
            for day in day_order:
                report_html += f"<td>{row[day]:.2f}</td>"
            report_html += "</tr>"

        report_html += "</table>"

        total_hours = df_filtered['Duration'].sum()
        report_html += f"<br><strong>Total Hours: {total_hours:.2f} hours</strong>"

        self.report_text_edit.setHtml(report_html)
