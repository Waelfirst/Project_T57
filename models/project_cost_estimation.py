# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import UserError
import base64
import io
import logging

_logger = logging.getLogger(__name__)


class ProjectCostEstimation(models.Model):
    _name = 'project.cost.estimation'
    _description = 'Project Cost Estimation'
    _inherit = ['mail.thread', 'mail.activity.mixin']

    name = fields.Char(string='Estimation Reference', required=True, copy=False, readonly=True,
                       default=lambda self: _('New'))
    project_id = fields.Many2one('project.definition', string='Project', required=True, ondelete='cascade')

    excel_file = fields.Binary(string='Estimation Excel', attachment=True)
    excel_filename = fields.Char(string='Filename', default='Cost_Estimation.xlsx')

    state = fields.Selection([
        ('draft', 'Draft'),
        ('in_progress', 'In Progress'),
        ('completed', 'Completed'),
        ('cancelled', 'Cancelled'),
    ], string='Status', default='draft', tracking=True)

    estimation_date = fields.Datetime(string='Estimation Date', default=fields.Datetime.now, tracking=True)
    last_update_date = fields.Datetime(string='Last Update', tracking=True)

    notes = fields.Text(string='Notes')
    company_id = fields.Many2one('res.company', string='Company', default=lambda self: self.env.company)

    @api.model
    def create(self, vals):
        if vals.get('name', _('New')) == _('New'):
            vals['name'] = self.env['ir.sequence'].next_by_code('project.cost.estimation') or _('New')
        return super(ProjectCostEstimation, self).create(vals)

    def action_generate_estimation_excel(self):
        """Generate Excel file with 10 editable weight columns"""
        self.ensure_one()

        if not self.project_id.product_line_ids:
            raise UserError(_('No products found in project to create estimation!'))

        try:
            import xlsxwriter
        except ImportError:
            raise UserError(_('Please install xlsxwriter library: pip install xlsxwriter'))

        try:
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})

            # Create main estimation sheet
            ws_estimation = workbook.add_worksheet('Cost Estimation')

            # Create instructions sheet
            ws_instructions = workbook.add_worksheet('Instructions')

            # Formats
            title_format = workbook.add_format({
                'bold': True,
                'font_size': 18,
                'font_color': '#1F4E78',
                'align': 'center',
                'valign': 'vcenter',
            })

            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4CAF50',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True,
            })

            locked_format = workbook.add_format({
                'bg_color': '#E8E8E8',
                'border': 1,
                'align': 'left',
                'locked': True,
            })

            unlocked_format = workbook.add_format({
                'bg_color': '#FFFFFF',
                'border': 1,
                'align': 'right',
                'num_format': '#,##0.00',
                'locked': False,
            })

            formula_format = workbook.add_format({
                'bg_color': '#FFF3CD',
                'border': 1,
                'align': 'right',
                'num_format': '#,##0.00',
                'locked': True,
            })

            # Add company header
            company = self.env.company
            row = 0

            if company.logo:
                try:
                    logo_data = base64.b64decode(company.logo)
                    image_data = io.BytesIO(logo_data)
                    ws_estimation.insert_image(row, 0, 'logo.png', {
                        'x_scale': 0.4,
                        'y_scale': 0.4,
                        'x_offset': 10,
                        'y_offset': 10,
                        'image_data': image_data,
                    })
                except Exception as e:
                    _logger.warning('Could not insert company logo: %s', str(e))

            ws_estimation.merge_range(row, 2, row, 6, company.name or 'Company Name', title_format)
            row += 2

            # Project information
            ws_estimation.merge_range(row, 0, row, 6, f'PROJECT COST ESTIMATION - {self.project_id.name}', title_format)
            row += 1

            ws_estimation.write(row, 0, 'Project:', locked_format)
            ws_estimation.merge_range(row, 1, row, 2, self.project_id.project_name, locked_format)
            ws_estimation.write(row, 3, 'Customer:', locked_format)
            ws_estimation.merge_range(row, 4, row, 6, self.project_id.partner_id.name, locked_format)
            row += 1

            ws_estimation.write(row, 0, 'Estimation Date:', locked_format)
            ws_estimation.merge_range(row, 1, row, 2,
                                      fields.Datetime.to_string(self.estimation_date), locked_format)
            ws_estimation.write(row, 3, 'Estimation Ref:', locked_format)
            ws_estimation.merge_range(row, 4, row, 6, self.name, locked_format)
            row += 2

            # Headers - NEW: 10 editable weight columns + Cost Price calculation
            headers = [
                'Product Code',
                'Product Name',
                'Quantity',
                'Unit Weight (kg)',
                'Weight 1\n(Editable)',
                'Weight 2\n(Editable)',
                'Weight 3\n(Editable)',
                'Weight 4\n(Editable)',
                'Weight 5\n(Editable)',
                'Weight 6\n(Editable)',
                'Weight 7\n(Editable)',
                'Weight 8\n(Editable)',
                'Weight 9\n(Editable)',
                'Weight 10\n(Editable)',
                'Total Weight',
                'Cost Price\n(Calculated)',
                'Sale Price\n(Editable)',
                'Total Cost',
                'Total Sale',
                'Profit',
                'Profit %'
            ]

            for col, header in enumerate(headers):
                ws_estimation.write(row, col, header, header_format)

            row += 1
            data_start_row = row

            # Add product lines
            for product_line in self.project_id.product_line_ids:
                col = 0
                # Product code (locked)
                ws_estimation.write(row, col, product_line.product_id.default_code or '', locked_format)
                col += 1

                # Product name (locked)
                ws_estimation.write(row, col, product_line.product_id.name or '', locked_format)
                col += 1

                # Quantity (locked)
                ws_estimation.write(row, col, product_line.quantity, locked_format)
                col += 1

                # Unit Weight (locked)
                ws_estimation.write(row, col, product_line.weight, locked_format)
                col += 1

                # 10 Editable Weight Columns (E to N) - columns 4-13
                for i in range(10):
                    ws_estimation.write(row, col, 0.0, unlocked_format)
                    col += 1

                # Total Weight (Formula) - Column O (14)
                # Sum of (Unit Weight * Weight1) + (Unit Weight * Weight2) + ... + (Unit Weight * Weight10)
                weight_formula = f'=D{row + 1}*(E{row + 1}+F{row + 1}+G{row + 1}+H{row + 1}+I{row + 1}+J{row + 1}+K{row + 1}+L{row + 1}+M{row + 1}+N{row + 1})'
                ws_estimation.write_formula(row, col, weight_formula, formula_format)
                col += 1

                # Cost Price (Calculated from Total Weight) - Column P (15)
                # Cost Price = Total Weight
                cost_formula = f'=O{row + 1}'
                ws_estimation.write_formula(row, col, cost_formula, formula_format)
                col += 1

                # Sale Price (EDITABLE) - Column Q (16)
                ws_estimation.write(row, col, product_line.sale_price, unlocked_format)
                col += 1

                # Total Cost (Formula) - Column R (17)
                # Total Cost = Quantity * Cost Price
                ws_estimation.write_formula(row, col, f'=C{row + 1}*P{row + 1}', formula_format)
                col += 1

                # Total Sale (Formula) - Column S (18)
                # Total Sale = Quantity * Sale Price
                ws_estimation.write_formula(row, col, f'=C{row + 1}*Q{row + 1}', formula_format)
                col += 1

                # Profit (Formula) - Column T (19)
                ws_estimation.write_formula(row, col, f'=S{row + 1}-R{row + 1}', formula_format)
                col += 1

                # Profit % (Formula) - Column U (20)
                ws_estimation.write_formula(row, col, f'=IF(S{row + 1}>0,T{row + 1}/S{row + 1}*100,0)', formula_format)
                col += 1

                row += 1

            # Totals row
            ws_estimation.write(row, 0, 'TOTAL', header_format)
            ws_estimation.write(row, 1, '', header_format)
            ws_estimation.write(row, 2, '', header_format)
            ws_estimation.write(row, 3, '', header_format)

            # Empty cells for weight columns
            for i in range(10):
                ws_estimation.write(row, 4 + i, '', header_format)

            ws_estimation.write(row, 14, '', header_format)  # Total Weight
            ws_estimation.write(row, 15, '', header_format)  # Cost Price
            ws_estimation.write(row, 16, '', header_format)  # Sale Price

            # Total Cost - Column R (17)
            ws_estimation.write_formula(row, 17, f'=SUM(R{data_start_row + 1}:R{row})', header_format)
            # Total Sale - Column S (18)
            ws_estimation.write_formula(row, 18, f'=SUM(S{data_start_row + 1}:S{row})', header_format)
            # Total Profit - Column T (19)
            ws_estimation.write_formula(row, 19, f'=SUM(T{data_start_row + 1}:T{row})', header_format)
            # Average Profit % - Column U (20)
            ws_estimation.write_formula(row, 20, f'=IF(S{row + 1}>0,T{row + 1}/S{row + 1}*100,0)', header_format)

            # Set column widths
            ws_estimation.set_column('A:A', 15)  # Product Code
            ws_estimation.set_column('B:B', 35)  # Product Name
            ws_estimation.set_column('C:C', 12)  # Quantity
            ws_estimation.set_column('D:D', 15)  # Unit Weight
            ws_estimation.set_column('E:N', 12)  # 10 Weight columns (editable)
            ws_estimation.set_column('O:O', 15)  # Total Weight
            ws_estimation.set_column('P:P', 15)  # Cost Price (calculated)
            ws_estimation.set_column('Q:Q', 15)  # Sale Price (editable)
            ws_estimation.set_column('R:U', 15)  # Total Cost, Total Sale, Profit, Profit %

            # Protect sheet (allow editing only weight columns and sale price)
            ws_estimation.protect('', {
                'objects': True,
                'scenarios': True,
                'format_cells': False,
                'format_columns': False,
                'format_rows': False,
            })

            # ==================== INSTRUCTIONS SHEET ====================
            self._create_instructions_sheet(ws_instructions, workbook)

            workbook.close()
            output.seek(0)

            file_data = base64.b64encode(output.read())
            filename = f'Cost_Estimation_{self.project_id.name.replace("/", "_")}.xlsx'

            self.write({
                'excel_file': file_data,
                'excel_filename': filename,
                'state': 'in_progress',
            })

            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': _('Success'),
                    'message': _('Cost estimation Excel file generated successfully! Download and edit it.'),
                    'type': 'success',
                    'sticky': False,
                }
            }

        except Exception as e:
            raise UserError(_('Error generating Excel file: %s') % str(e))

    def _create_instructions_sheet(self, worksheet, workbook):
        """Create instructions sheet"""
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 16,
            'font_color': '#1F4E78',
        })

        header_format = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'font_color': '#2E86C1',
        })

        text_format = workbook.add_format({
            'font_size': 11,
            'text_wrap': True,
            'valign': 'top',
        })

        warning_format = workbook.add_format({
            'font_size': 11,
            'font_color': '#D32F2F',
            'bold': True,
        })

        row = 0
        worksheet.write(row, 0, '📊 COST ESTIMATION - INSTRUCTIONS', title_format)
        row += 2

        instructions = [
            ('📝 How to Use:', header_format),
            ('1. Open the "Cost Estimation" sheet', text_format),
            ('2. Edit the editable columns (white background):', text_format),
            ('   • Weight 1 to Weight 10: Enter weight multipliers', text_format),
            ('   • Sale Price: Enter your desired sale price', text_format),
            ('3. Other columns are calculated automatically:', text_format),
            ('   • Total Weight = Unit Weight × (Weight1 + Weight2 + ... + Weight10)', text_format),
            ('   • Cost Price = Total Weight', text_format),
            ('   • Total Cost = Quantity × Cost Price', text_format),
            ('   • Total Sale = Quantity × Sale Price', text_format),
            ('   • Profit = Total Sale - Total Cost', text_format),
            ('   • Profit % = (Profit / Total Sale) × 100', text_format),
            ('4. Save the file after editing', text_format),
            ('5. Upload the edited file back to Odoo using "Import Updated Prices" button', text_format),
            ('', text_format),
            ('⚠️ IMPORTANT NOTES:', warning_format),
            ('• Do NOT modify product names or codes', text_format),
            ('• Do NOT add or delete rows', text_format),
            ('• Do NOT modify quantities or unit weights', text_format),
            ('• Only edit Weight 1-10 and Sale Price columns (white cells)', text_format),
            ('• Yellow cells contain formulas - do not edit them', text_format),
            ('• Save in Excel format (.xlsx)', text_format),
            ('', text_format),
            ('📋 Column Descriptions:', header_format),
            ('• Product Code: Internal product reference (locked)', text_format),
            ('• Product Name: Product description (locked)', text_format),
            ('• Quantity: Required quantity (locked)', text_format),
            ('• Unit Weight: Weight per unit in kg (locked)', text_format),
            ('• Weight 1-10: Weight multipliers (EDITABLE - Enter values)', text_format),
            ('• Total Weight: Automatically calculated = Unit Weight × Sum(Weight1-10)', text_format),
            ('• Cost Price: Automatically calculated = Total Weight', text_format),
            ('• Sale Price: Unit sale price (EDITABLE)', text_format),
            ('• Total Cost: Automatically calculated (Quantity × Cost Price)', text_format),
            ('• Total Sale: Automatically calculated (Quantity × Sale Price)', text_format),
            ('• Profit: Automatically calculated (Total Sale - Total Cost)', text_format),
            ('• Profit %: Automatically calculated (Profit / Total Sale × 100)', text_format),
            ('', text_format),
            ('💡 Example:', header_format),
            ('If Unit Weight = 10 kg, and you enter:', text_format),
            ('  Weight1 = 2, Weight2 = 3, Weight3 = 1, others = 0', text_format),
            ('Then: Total Weight = 10 × (2+3+1) = 60 kg', text_format),
            ('And: Cost Price = 60', text_format),
        ]

        for instruction, fmt in instructions:
            worksheet.write(row, 0, instruction, fmt)
            row += 1

        worksheet.set_column('A:A', 80)

    def action_download_estimation(self):
        """Download the estimation Excel file"""
        self.ensure_one()

        if not self.excel_file:
            return self.action_generate_estimation_excel()

        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content?model={self._name}&id={self.id}&field=excel_file&filename_field=excel_filename&download=true',
            'target': 'new',
        }

    def action_import_updated_prices(self):
        """Open wizard to import updated prices from Excel"""
        self.ensure_one()

        return {
            'type': 'ir.actions.act_window',
            'name': _('Import Updated Prices'),
            'res_model': 'project.estimation.import.wizard',
            'view_mode': 'form',
            'target': 'new',
            'context': {
                'default_estimation_id': self.id,
                'default_project_id': self.project_id.id,
            }
        }

    def action_complete(self):
        """Mark estimation as completed"""
        self.write({
            'state': 'completed',
            'last_update_date': fields.Datetime.now(),
        })

    def action_cancel(self):
        """Cancel estimation"""
        self.write({'state': 'cancelled'})

    def action_reset_to_draft(self):
        """Reset to draft"""
        self.write({'state': 'draft'})