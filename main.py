import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
from fpdf import FPDF
import smtplib
from email.message import EmailMessage
import os

class Tickets:
    
    def __init__(self, ticket_id, date_opened, date_closed, category, status):
        self.ticket_id = ticket_id
        self.date_opened = pd.to_datetime(date_opened)
        self.date_closed = pd.to_datetime(date_closed) if pd.notna(date_closed) else None
        self.category = category
        self.status= status.lower()
        
class TicketReport:
    
    def __init__(self, tickets):
        self.tickets = tickets
        self.df = pd.DataFrame([t.__dict__ for t in tickets])
    
    def tickets_per_day(self):
        return self.df.groupby(self.df['date_opened'].dt.date)['ticket_id'].count()
        
    def closed_tickets_per_day(self):
        closed_df = self.df[self.df['status'] == 'closed']
        return closed_df.groupby(closed_df['date_closed'].dt.date)['ticket_id'].count()
        
    def category_distribution(self):
        return self.df['category'].value_counts()
        
    def closure_rate(self):
        total = len(self.df)
        closed = len(self.df[self.df['status'] == 'closed'])
        return round((closed / total) * 100, 2) if total else 0
        
    def summary(self):
        return{
            "Total Tickets": len(self.df),
            "Closed Tickets": len(self.df[self.df['status'] == 'closed']),
            "Closure Rate (%)": self.closure_rate(),
            "Tickets per Day": self.tickets_per_day().to_dict(),
            "Closed per Day": self.closed_tickets_per_day().to_dict(),
            "By Category": self.category_distribution().to_dict()
        }
        
    def export_to_excel(self, path="helpdesk_report.xlsx"):
        summary = self.summary()
        with pd.ExcelWriter(path, engine='openpyxl') as writer:
            self.df.to_excel(writer, sheet_name='Tickets', index=False)
            pd.DataFrame(summary["Tickets per Day"].items(), columns=['Date', 'Opened']).to_excel(writer, sheet_name='Opened per Day', index=False)
            pd.DataFrame(summary["Closed per Day"].items(), columns=['Date', 'Closed']).to_excel(writer, sheet_name='Closed per Day', index=False)
            pd.DataFrame(summary["By Category"].items(), columns=['Category', 'Count']).to_excel(writer, sheet_name='By Category', index=False)
        return path
        
    def export_to_pdf(self, path="helpdesk_report.pdf"):
            """Export summary as a formatted PDF report."""
            summary = self.summary()
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(0, 10, "Helpdesk Daily Report", ln=True, align="C")
            pdf.ln(10)
            pdf.set_font("Arial", size=12)
    
            for key, value in summary.items():
                if isinstance(value, dict):
                    pdf.cell(0, 10, f"{key}:", ln=True)
                    for sub_k, sub_v in value.items():
                        pdf.cell(0, 10, f"   {sub_k}: {sub_v}", ln=True)
                else:
                    pdf.cell(0, 10, f"{key}: {value}", ln=True)
                pdf.ln(5)
    
            pdf.output(path)
            return path
        
class ReportGenarator:
    def __init__(self, excel_path, email_config=None):
        self.excel_path = excel_path
        self.email_config = email_config or {}
        
    def load_tickets(self):
        df = pd.read_excel(self.excel_path)
        tickets = [
            Ticket(
                row['Ticket_ID'],
                row['Date_Opened'],
                row['Date_Closed'],
                row['Category'],
                row['Status']
            )
            for _, row in df.iterrows()
        ]
        return tickets
        
    def generate_report(self):
        tickets = self.load_tickets()
        report = TicketReport(tickets)
        summary = report.summary()

        print("\n=== Helpdesk Report Summary ===")
        for k, v in summary.items():
            print(f"{k}: {v}")

        # Export reports
        excel_path = report.export_to_excel()
        pdf_path = report.export_to_pdf()

        print(f"\n✅ Excel exported to: {excel_path}")
        print(f"✅ PDF exported to: {pdf_path}")

        # Optionally send email
        if self.email_config:
            self.send_email([excel_path, pdf_path])
        
    def send_email(self, attachments):
        """Send the report by email."""
        cfg = self.email_config
        msg = EmailMessage()
        msg['Subject'] = cfg.get('subject', 'Helpdesk Report')
        msg['From'] = cfg['from']
        msg['To'] = cfg['to']
        msg.set_content(cfg.get('body', 'Please find attached the daily Helpdesk report.'))

        for file in attachments:
            with open(file, 'rb') as f:
                msg.add_attachment(
                    f.read(),
                    maintype='application',
                    subtype='octet-stream',
                    filename=os.path.basename(file)
                )

        with smtplib.SMTP_SSL(cfg['smtp_server'], cfg['smtp_port']) as smtp:
            smtp.login(cfg['from'], cfg['password'])
            smtp.send_message(msg)

        print(f"📧 Email sent to {cfg['to']} with attachments.")


if __name__ == "__main__":
    # Example usage
    email_settings = {
        "from": "your_email@gmail.com",
        "password": "your_app_password",  # Use app password, not your real one
        "to": "manager@company.com",
        "smtp_server": "smtp.gmail.com",
        "smtp_port": 465,
        "subject": "Daily Helpdesk Report",
        "body": "Attached are today's Helpdesk reports (Excel + PDF)."
    }

    generator = ReportGenerator("tickets.xlsx", email_config=email_settings)
    generator.generate_report()
        
