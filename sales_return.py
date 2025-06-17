from odoo import models, fields
from datetime import datetime
import xlsxwriter
from io import BytesIO
import base64

class SalesReturn(models.TransientModel):
    _name = 'sales.return'
    _description = 'Sales Return'

    def GetInput(self):
        month = fields.Selection([
            ('1', 'January'), ('2', 'February'), ('3', 'March'), ('4', 'April'),
            ('5', 'May'), ('6', 'June'), ('7', 'July'), ('8', 'August'),
            ('9', 'September'), ('10', 'October'), ('11', 'November'), ('12', 'December')
        ], string='Enter Month:', required=True)

        year = fields.Selection([
            ('10', '2030'), ('9', '2029'), ('8', '2028'), ('7', '2027'),
            ('6', '2026'), ('0', '2025'), ('1', '2024'), ('2', '2023'),
            ('3', '2022'), ('4', '2021'), ('5', '2020')
        ], string="Enter Year:", required=True)

        previous_month_count = fields.Integer(string="Enter Number of Previous Month:", required=True)

        return month,year,previous_month_count
    
    file_name = fields.Html("Generated Report", sanitize=False)

    def generate_excel(self):
        # Year mapping
        year_mapping = {
            '10': 2030, '9': 2029, '8': 2028, '7': 2027, '6': 2026,
            '0': 2025, '1': 2024, '2': 2023, '3': 2022, '4': 2021, '5': 2020
        }
        year = year_mapping.get(self.year)
        month_number = int(self.month)
        previous_month_count = self.previous_month_count

        month_list = {
            1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 6: 'June',
            7: 'July', 8: 'August', 9: 'September', 10: 'October', 11: 'November', 12: 'December'
        }

        # Get current company address dynamically
        company = self.env.company
        company_address = filter(None, [
            company.name,
            company.street,
            f"{company.zip or ''} - {company.city or ''}",
            f"{company.state_id.name or ''} ({company.country_id.code or ''})" if company.state_id else company.country_id.name,
            company.country_id.name,
        ])
        address_lines = list(company_address)

        # Create Excel
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet()

        # Styles
        header_format = workbook.add_format({'bold': True, 'bg_color': "#AAAAAA", 'border': 1})
        merge_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})

        # Dynamic company info header
        row = 0
        for line in address_lines:
            worksheet.merge_range(row, 2, row, 6, line, merge_format)
            row += 1
        worksheet.merge_range(row, 2, row, 6, 'Company Report', merge_format)
        row += 1

        # Table header
        worksheet.write(row, 0, "S.No", header_format)
        worksheet.write(row, 1, "Customer Name", header_format)

        col = 2
        current_month = month_number
        current_year = year
        months = []

        for _ in range(previous_month_count + 1):
            if current_month < 1:
                current_month = 12
                current_year -= 1
            months.append((current_month, current_year))
            current_month -= 1

        months.reverse()

        for m, y in months:
            worksheet.write(row, col, f"{month_list[m]}, {y}", header_format)
            col += 1

        # Column width adjustments
        worksheet.set_column('A:A', 10)
        worksheet.set_column('B:B', 20)
        worksheet.set_column('C:Z', 18)

        # Finalize Excel
        workbook.close()
        output.seek(0)
        file_content = output.read()

        filename = f"BeforeMonth_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        attachment = self.env['ir.attachment'].create({
            'name': filename,
            'type': 'binary',
            'datas': base64.b64encode(file_content),
            'res_model': self._name,
            'res_id': self.id,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })

        download_url = f"/web/content/{attachment.id}?download=true"
        self.file_name = f'<a href="{download_url}" target="_blank">{filename}</a>'

        return {
            'type': 'ir.actions.act_window',
            'res_model': 'sales.return',
            'res_id': self.id,
            'view_mode': 'form',
            'target': 'new',
            'context': dict(self.env.context),
        }
