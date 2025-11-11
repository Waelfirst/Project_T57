# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import UserError
import xlsxwriter
import io
import base64


class TemplateGeneratorWizard(models.TransientModel):
    _name = 'template.generator.wizard'
    _description = 'Excel Template Generator'

    template_type = fields.Selection([
        ('components', 'Components Template'),
        ('bom_materials', 'BOM Materials Template'),
        ('bom_operations', 'BOM Operations Template'),
        ('complete', 'Complete Import Template (All Sheets)'),
        ('all_separate', 'All Templates (Separate Files)'),
    ], string='Template Type', default='complete', required=True)

    def action_generate_template(self):
        """Generate and download the selected template"""
        self.ensure_one()

        try:
            import xlsxwriter
        except ImportError:
            raise UserError(_('xlsxwriter library not installed. Please install it: pip install xlsxwriter'))

        if self.template_type == 'all_separate':
            return self._generate_all_templates()
        else:
            return self._generate_single_template()

    def _generate_single_template(self):
        """Generate a single template file"""
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})

        if self.template_type == 'components':
            self._create_components_sheet(workbook)
            filename = 'Components_Import_Template.xlsx'
        elif self.template_type == 'bom_materials':
            self._create_materials_sheet(workbook)
            filename = 'BOM_Materials_Import_Template.xlsx'
        elif self.template_type == 'bom_operations':
            self._create_operations_sheet(workbook)
            filename = 'BOM_Operations_Import_Template.xlsx'
        else:  # complete
            self._create_components_sheet(workbook)
            self._create_materials_sheet(workbook)
            self._create_operations_sheet(workbook)
            filename = 'Complete_Import_Template.xlsx'

        workbook.close()
        output.seek(0)

        file_data = base64.b64encode(output.read())

        attachment = self.env['ir.attachment'].create({
            'name': filename,
            'datas': file_data,
            'res_model': self._name,
            'res_id': self.id,
            'type': 'binary',
        })

        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/%s?download=true' % attachment.id,
            'target': 'new',
        }

    def _generate_all_templates(self):
        """Generate all templates as separate files and return as zip"""
        raise UserError(_('Generate individual templates one at a time, or use "Complete Import Template" to get all sheets in one file.'))

    def _create_components_sheet(self, workbook):
        """Create Components sheet"""
        worksheet = workbook.add_worksheet('Components')

        # Formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4CAF50',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_size': 11,
        })

        example_format = workbook.add_format({
            'bg_color': '#E8F5E9',
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
        })

        instruction_format = workbook.add_format({
            'bg_color': '#FFF9C4',
            'border': 1,
            'text_wrap': True,
            'valign': 'top',
            'font_size': 9,
        })

        # Set column widths
        worksheet.set_column('A:A', 35)
        worksheet.set_column('B:B', 12)
        worksheet.set_column('C:C', 15)
        worksheet.set_column('D:D', 15)
        worksheet.set_column('E:E', 20)

        # Headers
        headers = ['Component Name*', 'Quantity*', 'Weight (kg)', 'Cost Price*', 'BOM Code']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        # Instructions row
        instructions = [
            'Product name or internal reference\n(Must exist in Odoo)',
            'Numeric quantity\nrequired',
            'Weight in kg\n(numeric)',
            'Unit cost price\n(numeric)',
            'Code to link with BOM\n(optional, text)'
        ]
        for col, instruction in enumerate(instructions):
            worksheet.write(1, col, instruction, instruction_format)

        worksheet.set_row(1, 45)

        # Example data
        examples = [
            ['Steel Sheet 2mm', 2, 5.5, 50.00, 'BOM-001'],
            ['Aluminum Plate 3mm', 4, 3.2, 75.00, 'BOM-002'],
            ['Plastic Housing ABS', 1, 0.8, 25.00, 'BOM-003'],
            ['Screws M6 Stainless', 10, 0.05, 0.50, ''],
            ['Motor Assembly 12V', 1, 2.0, 150.00, 'BOM-004'],
        ]

        for row_idx, example in enumerate(examples, start=2):
            for col_idx, value in enumerate(example):
                worksheet.write(row_idx, col_idx, value, example_format)

    def _create_materials_sheet(self, workbook):
        """Create BOM Materials sheet"""
        worksheet = workbook.add_worksheet('BOM Materials')

        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#2196F3',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_size': 11,
        })

        example_format = workbook.add_format({
            'bg_color': '#E3F2FD',
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
        })

        instruction_format = workbook.add_format({
            'bg_color': '#FFF9C4',
            'border': 1,
            'text_wrap': True,
            'valign': 'top',
            'font_size': 9,
        })

        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 35)
        worksheet.set_column('C:C', 15)
        worksheet.set_column('D:D', 15)

        headers = ['BOM Code*', 'Material Name*', 'Quantity*', 'Unit']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        instructions = [
            'Must match BOM Code\nfrom Components',
            'Raw material product name\n(Must exist in Odoo)',
            'Quantity needed\n(numeric)',
            'Unit of measure\n(kg, pcs, meters, etc.)'
        ]
        for col, instruction in enumerate(instructions):
            worksheet.write(1, col, instruction, instruction_format)

        worksheet.set_row(1, 45)

        examples = [
            ['BOM-001', 'Steel Raw Material Grade A', 6, 'kg'],
            ['BOM-001', 'Coating Material Epoxy', 0.5, 'kg'],
            ['BOM-002', 'Aluminum Sheet 6061', 5, 'kg'],
            ['BOM-002', 'Anodizing Chemical', 0.3, 'liter'],
            ['BOM-003', 'Plastic Pellets ABS', 1, 'kg'],
            ['BOM-003', 'Paint Black RAL9005', 0.1, 'liter'],
        ]

        for row_idx, example in enumerate(examples, start=2):
            for col_idx, value in enumerate(example):
                worksheet.write(row_idx, col_idx, value, example_format)

    def _create_operations_sheet(self, workbook):
        """Create BOM Operations sheet"""
        worksheet = workbook.add_worksheet('BOM Operations')

        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#FF9800',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_size': 11,
        })

        example_format = workbook.add_format({
            'bg_color': '#FFF3E0',
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
        })

        instruction_format = workbook.add_format({
            'bg_color': '#FFF9C4',
            'border': 1,
            'text_wrap': True,
            'valign': 'top',
            'font_size': 9,
        })

        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 30)
        worksheet.set_column('C:C', 25)
        worksheet.set_column('D:D', 20)

        headers = ['BOM Code*', 'Operation Name*', 'Workcenter', 'Duration (minutes)*']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        instructions = [
            'Must match BOM Code\nfrom Components',
            'Name of manufacturing\noperation',
            'Workcenter name\n(Must exist in Odoo)',
            'Time in minutes\n(numeric)'
        ]
        for col, instruction in enumerate(instructions):
            worksheet.write(1, col, instruction, instruction_format)

        worksheet.set_row(1, 45)

        examples = [
            ['BOM-001', 'Cutting', 'CNC Machine 1', 15],
            ['BOM-001', 'Bending', 'Press Machine', 10],
            ['BOM-001', 'Coating', 'Coating Line', 30],
            ['BOM-002', 'Material Cutting', 'CNC Machine 2', 12],
            ['BOM-002', 'Drilling', 'Drill Press', 8],
            ['BOM-003', 'Injection Molding', 'Molding Machine 1', 5],
            ['BOM-003', 'Painting', 'Paint Booth', 10],
        ]

        for row_idx, example in enumerate(examples, start=2):
            for col_idx, value in enumerate(example):
                worksheet.write(row_idx, col_idx, value, example_format)