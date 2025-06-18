from odoo import models, fields, api
from datetime import datetime
from dateutil.relativedelta import relativedelta
import xlsxwriter
from io import BytesIO
import base64


class SalesReturn(models.TransientModel):
    _name = 'sales.return'
    _description = 'Sales Return'

    month = fields.Selection([
        ('1', 'January'), ('2', 'February'), ('3', 'March'), ('4', 'April'),
        ('5', 'May'), ('6', 'June'), ('7', 'July'), ('8', 'August'),
        ('9', 'September'), ('10', 'October'), ('11', 'November'), ('12', 'December')
    ], string='Enter Month:', required=True)

    year = fields.Selection([(str(y), str(y)) for y in range(2020, 2031)], string="Enter Year:", required=True)
    previous_month_count = fields.Integer(string="Enter Number of Previous Months:", required=True)

    file_name = fields.Char("Generated Report Name", readonly=True)
    file_link = fields.Html("Download File", compute="_compute_file_link", sanitize=False)

    @api.depends('file_name')
    def _compute_file_link(self):
        for rec in self:
            rec.file_link = ''  # Always assign a default value first
            if rec.file_name:
                attachment = self.env['ir.attachment'].search([
                    ('res_model', '=', rec._name),
                    ('res_id', '=', rec.id),
                    ('name', '=', rec.file_name)
                ], limit=1)
                if attachment:
                    rec.file_link = f'<a href="/web/content/{attachment.id}?download=true" target="_blank">{rec.file_name}</a>'

    def generate_excel(self):
        
        month_number = int(self.month)
        year = int(self.year)
        prev_months = self.previous_month_count

        month_list = {
            1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 6: 'June',
            7: 'July', 8: 'August', 9: 'September', 10: 'October', 11: 'November', 12: 'December'
        }

        # Build list of (month, year) for report
        start_date = datetime(year, month_number, 1)
        months = []
        for i in range(prev_months + 1):
            dt = start_date - relativedelta(months=i)
            months.append((dt.month, dt.year))
        months.reverse()

        # Address header
        company = self.env.company
        address_lines = list(filter(None, [
            company.name,
            company.street,
            f"{company.zip or ''} - {company.city or ''}"
        ]))

        # Create Excel
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet("Sales Return Report")

        header_format = workbook.add_format({'bold': True, 'bg_color': "#D3D3D3",'align': 'center', 'border': 1})
        merge_format = workbook.add_format({'bold': True, 'bg_color': "#FFF064", 'align': 'center', 'valign': 'vcenter', 'border': 1})

        row = 0
        for line in address_lines:
            worksheet.merge_range(row, 2, row, 6, line, merge_format)
            row += 1

        worksheet.write(row, 0, "S.No", header_format)
        worksheet.write(row, 1, "Customer Name", header_format)

        col = 2
        for m, y in months:
            worksheet.write(row, col, f"{month_list[m]} - {y}", header_format)
            col += 1

        # Fetch sales report
        row += 1
        serial = 1
        customers = self.env['res.partner'].search([('customer_rank', '>', 0)], order='name')
        for customer in customers:
            worksheet.write(row, 0, serial)
            worksheet.write(row, 1, customer.name)
            col = 2
            for m, y in months:
                date_start = datetime(y, m, 1)
                date_end = date_start + relativedelta(months=1, days=-1)
                orders = self.env['sale.order'].search([
                    ('partner_id', '=', customer.id),
                    ('date_order', '>=', date_start),
                    ('date_order', '<=', date_end),
                    ('state', 'in', ['sale', 'done'])
                ])
                '''if not orders:
                    continue'''
                total = sum(orders.mapped('amount_total'))
                worksheet.write(row, col, total)
                col += 1
            row += 1
            serial += 1

        worksheet.set_column('A:A', 8)
        worksheet.set_column('B:B', 25)
        worksheet.set_column('C:Z', 18)

        workbook.close()
        output.seek(0)
        file_content = output.read()

        filename = f"Customer_Sales_Return_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        attachment = self.env['ir.attachment'].create({
            'name': filename,
            'type': 'binary',
            'datas': base64.b64encode(file_content),
            'res_model': self._name,
            'res_id': self.id,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })

        self.file_name = filename

        return {
            'type': 'ir.actions.act_window',
            'res_model': 'sales.return',
            'res_id': self.id,
            'view_mode': 'form',
            'target': 'new',
            'context': dict(self.env.context),
        }
    

