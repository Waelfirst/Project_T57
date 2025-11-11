# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import ValidationError, UserError
import base64
import io


class ExcelExportHelper:
    """Helper class for Excel exports with company header"""

    @staticmethod
    def add_company_header(workbook, worksheet, env, title="Export Report", start_row=0):
        """
        Add company header with logo to worksheet

        Args:
            workbook: xlsxwriter workbook object
            worksheet: xlsxwriter worksheet object
            env: Odoo environment
            title: Title to display in header
            start_row: Starting row for header (default: 0)

        Returns:
            int: Next available row after header
        """
        company = env.company

        # Formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 18,
            'font_color': '#1F4E78',
            'align': 'left',
            'valign': 'vcenter',
        })

        company_format = workbook.add_format({
            'font_size': 12,
            'font_color': '#1F4E78',
            'align': 'left',
            'valign': 'vcenter',
        })

        info_format = workbook.add_format({
            'font_size': 10,
            'font_color': '#666666',
            'align': 'left',
            'valign': 'vcenter',
        })

        border_format = workbook.add_format({
            'bottom': 2,
            'bottom_color': '#4CAF50',
        })

        current_row = start_row

        # Add logo if available
        logo_col = 0
        info_col = 2

        if company.logo:
            try:
                # Decode logo
                logo_data = base64.b64decode(company.logo)

                # Insert logo directly from bytes (no temp file needed)
                # FIXED: Use BytesIO to pass image data directly instead of temp file
                image_data = io.BytesIO(logo_data)
                worksheet.insert_image(current_row, logo_col, 'logo.png', {
                    'x_scale': 0.5,
                    'y_scale': 0.5,
                    'x_offset': 10,
                    'y_offset': 10,
                    'image_data': image_data,
                })

            except Exception as e:
                # If logo fails, continue without it
                import logging
                _logger = logging.getLogger(__name__)
                _logger.warning('Could not insert company logo: %s', str(e))

        # Add company information
        worksheet.merge_range(current_row, info_col, current_row, info_col + 3,
                              company.name or 'Company Name', company_format)
        current_row += 1

        # Add address if available
        address_parts = []
        if company.street:
            address_parts.append(company.street)
        if company.street2:
            address_parts.append(company.street2)
        if company.city:
            address_parts.append(company.city)
        if company.zip:
            address_parts.append(company.zip)
        if company.country_id:
            address_parts.append(company.country_id.name)

        if address_parts:
            address = ', '.join(address_parts)
            worksheet.merge_range(current_row, info_col, current_row, info_col + 3,
                                  address, info_format)
            current_row += 1

        # Add contact info
        contact_parts = []
        if company.phone:
            contact_parts.append(f"Tel: {company.phone}")
        if company.email:
            contact_parts.append(f"Email: {company.email}")
        if company.website:
            contact_parts.append(f"Web: {company.website}")

        if contact_parts:
            contact = ' | '.join(contact_parts)
            worksheet.merge_range(current_row, info_col, current_row, info_col + 3,
                                  contact, info_format)
            current_row += 1

        # Add title
        current_row += 1
        worksheet.merge_range(current_row, 0, current_row, info_col + 3,
                              title, title_format)
        current_row += 1

        # Add export date
        from datetime import datetime
        export_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        worksheet.merge_range(current_row, 0, current_row, info_col + 3,
                              f"Export Date: {export_date}", info_format)
        current_row += 1

        # Add separator line
        worksheet.merge_range(current_row, 0, current_row, info_col + 3,
                              '', border_format)
        current_row += 2  # Extra space after header

        return current_row


class ProjectProductPricing(models.Model):
    _name = 'project.product.pricing'
    _description = 'Project Product Pricing'
    _inherit = ['mail.thread', 'mail.activity.mixin']
    _order = 'pricing_date desc, version desc'

    name = fields.Char(
        string='Pricing Code',
        required=True,
        copy=False,
        readonly=True,
        default=lambda self: _('New'),
        tracking=True
    )
    pricing_date = fields.Date(
        string='Pricing Date',
        required=True,
        default=fields.Date.context_today,
        tracking=True
    )
    version = fields.Integer(
        string='Version Number',
        required=True,
        default=1,
        tracking=True
    )
    partner_id = fields.Many2one(
        'res.partner',
        string='Customer',
        required=True,
        domain=[('customer_rank', '>', 0)],
        tracking=True
    )
    project_id = fields.Many2one(
        'project.definition',
        string='Project',
        required=True,
        tracking=True
    )
    product_id = fields.Many2one(
        'product.product',
        string='Product',
        required=True,
        tracking=True
    )
    quantity = fields.Float(
        string='Product Quantity',
        digits='Product Unit of Measure',
        compute='_compute_product_data',
        store=True,
        readonly=False
    )
    weight = fields.Float(
        string='Product Weight',
        digits='Stock Weight',
        compute='_compute_product_data',
        store=True,
        readonly=False
    )
    state = fields.Selection([
        ('draft', 'Draft'),
        ('confirmed', 'Confirmed'),
        ('approved', 'Approved'),
        ('cancelled', 'Cancelled'),
    ], string='Status', default='draft', tracking=True)

    component_line_ids = fields.One2many(
        'project.product.component',
        'pricing_id',
        string='Product Components'
    )

    total_component_cost = fields.Float(
        string='Total Component Cost',
        compute='_compute_totals',
        store=True,
        digits='Product Price'
    )

    company_id = fields.Many2one(
        'res.company',
        string='Company',
        default=lambda self: self.env.company
    )

    notes = fields.Text(string='Notes')

    @api.model
    def create(self, vals):
        if vals.get('name', _('New')) == _('New'):
            vals['name'] = self.env['ir.sequence'].next_by_code('project.product.pricing') or _('New')
        return super(ProjectProductPricing, self).create(vals)

    @api.depends('project_id', 'product_id')
    def _compute_product_data(self):
        for record in self:
            if record.product_id and record.project_id:
                product_line = record.project_id.product_line_ids.filtered(
                    lambda l: l.product_id == record.product_id
                )
                if product_line:
                    record.quantity = product_line[0].quantity
                    record.weight = product_line[0].weight
                else:
                    record.quantity = 0.0
                    record.weight = 0.0
            else:
                record.quantity = 0.0
                record.weight = 0.0

    @api.model
    def create(self, vals):
        """Trigger project state update when pricing is created"""
        if vals.get('name', _('New')) == _('New'):
            vals['name'] = self.env['ir.sequence'].next_by_code('project.product.pricing') or _('New')

        pricing = super(ProjectProductPricing, self).create(vals)

        # Update project state
        if pricing.project_id and pricing.project_id.auto_update_state:
            pricing.project_id.update_project_state()

        return pricing

    def write(self, vals):
        """Trigger project state update when pricing state changes"""
        result = super(ProjectProductPricing, self).write(vals)

        # Update project state if state changed
        if 'state' in vals:
            for pricing in self:
                if pricing.project_id and pricing.project_id.auto_update_state:
                    pricing.project_id.update_project_state()

        return result

    def action_confirm(self):
        """Update project state when confirming"""
        self.write({'state': 'confirmed'})

        # Trigger project state update
        for pricing in self:
            if pricing.project_id and pricing.project_id.auto_update_state:
                pricing.project_id.update_project_state()

    def action_approve(self):
        """Update project state when approving"""
        self.write({'state': 'approved'})

        # Trigger project state update
        for pricing in self:
            if pricing.project_id and pricing.project_id.auto_update_state:
                pricing.project_id.update_project_state()
    @api.depends('component_line_ids.total_cost')
    def _compute_totals(self):
        for record in self:
            record.total_component_cost = sum(line.total_cost for line in record.component_line_ids)

    @api.onchange('partner_id')
    def _onchange_partner_id(self):
        self.project_id = False
        self.product_id = False
        if self.partner_id:
            return {'domain': {'project_id': [('partner_id', '=', self.partner_id.id)]}}
        return {'domain': {'project_id': []}}

    @api.onchange('project_id')
    def _onchange_project_id(self):
        self.product_id = False
        if self.project_id:
            product_ids = self.project_id.product_line_ids.mapped('product_id').ids
            return {'domain': {'product_id': [('id', 'in', product_ids)]}}
        return {'domain': {'product_id': []}}

    def action_confirm(self):
        self.write({'state': 'confirmed'})

    def action_approve(self):
        self.write({'state': 'approved'})

    def action_cancel(self):
        self.write({'state': 'cancelled'})

    def action_draft(self):
        self.write({'state': 'draft'})

    def action_create_new_version(self):
        self.ensure_one()
        new_version = self.copy({
            'version': self.version + 1,
            'pricing_date': fields.Date.context_today(self),
            'state': 'draft',
        })
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'project.product.pricing',
            'view_mode': 'form',
            'res_id': new_version.id,
            'target': 'current',
        }

    def action_import_components(self):
        """Import components only"""
        self.ensure_one()
        return {
            'name': _('Import Components from Excel'),
            'type': 'ir.actions.act_window',
            'res_model': 'import.components.wizard',
            'view_mode': 'form',
            'target': 'new',
            'context': {
                'default_pricing_id': self.id,
                'default_product_id': self.product_id.id,
                'default_import_mode': 'components',
            }
        }

    def action_import_bom_materials(self):
        """Import BOM materials only"""
        self.ensure_one()
        return {
            'name': _('Import BOM Materials from Excel'),
            'type': 'ir.actions.act_window',
            'res_model': 'import.components.wizard',
            'view_mode': 'form',
            'target': 'new',
            'context': {
                'default_pricing_id': self.id,
                'default_product_id': self.product_id.id,
                'default_import_mode': 'bom_materials',
            }
        }

    def action_import_bom_operations(self):
        """Import BOM operations only"""
        self.ensure_one()
        return {
            'name': _('Import BOM Operations from Excel'),
            'type': 'ir.actions.act_window',
            'res_model': 'import.components.wizard',
            'view_mode': 'form',
            'target': 'new',
            'context': {
                'default_pricing_id': self.id,
                'default_product_id': self.product_id.id,
                'default_import_mode': 'bom_operations',
            }
        }

    def action_export_components_excel(self):
        """Export components with specifications, material breakdown and weight comparison"""
        self.ensure_one()

        if not self.component_line_ids:
            raise UserError(_('No components to export!'))

        try:
            import xlsxwriter

            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})

            # Create worksheets
            ws_summary = workbook.add_worksheet('Summary')
            ws_components = workbook.add_worksheet('Components Detail')
            ws_specifications = workbook.add_worksheet('Specifications')
            ws_materials = workbook.add_worksheet('Materials Breakdown')

            # Formats
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4CAF50',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            title_format = workbook.add_format({
                'bold': True,
                'font_size': 14,
                'bg_color': '#2E86C1',
                'font_color': 'white',
                'border': 1,
            })

            cell_format = workbook.add_format({
                'border': 1,
                'align': 'left',
                'valign': 'vcenter'
            })

            number_format = workbook.add_format({
                'border': 1,
                'align': 'right',
                'num_format': '0.00'
            })

            currency_format = workbook.add_format({
                'border': 1,
                'align': 'right',
                'num_format': '#,##0.00'
            })

            text_wrap_format = workbook.add_format({
                'border': 1,
                'align': 'left',
                'valign': 'top',
                'text_wrap': True
            })

            # ==================== SUMMARY SHEET WITH HEADER ====================
            row = ExcelExportHelper.add_company_header(
                workbook, ws_summary, self.env,
                title=f"Product Pricing Summary - {self.name}",
                start_row=0
            )

            summary_data = [
                ['Pricing Code:', self.name],
                ['Customer:', self.partner_id.name],
                ['Project:', self.project_id.name],
                ['Product:', self.product_id.display_name],
                ['Product Quantity:', self.quantity],
                ['Product Unit Weight (kg):', self.product_id.weight],
                ['Product Total Weight (kg):', self.weight],
                ['Pricing Date:', self.pricing_date.strftime('%Y-%m-%d') if self.pricing_date else ''],
                ['Version:', self.version],
                ['Status:', dict(self._fields['state'].selection).get(self.state)],
            ]

            for label, value in summary_data:
                ws_summary.write(row, 0, label, cell_format)
                if isinstance(value, (int, float)):
                    ws_summary.write(row, 1, value, number_format)
                else:
                    ws_summary.write(row, 1, value, cell_format)
                row += 1

            row += 2

            total_component_weight = sum(c.weight * c.quantity for c in self.component_line_ids)
            total_component_cost = self.total_component_cost

            total_material_weight = 0
            total_material_cost = 0
            for comp in self.component_line_ids:
                if comp.bom_id:
                    for bom_line in comp.bom_id.bom_line_ids:
                        material = bom_line.product_id
                        material_qty = bom_line.product_qty * comp.quantity
                        total_material_weight += material.weight * material_qty
                        total_material_cost += material.standard_price * material_qty

            ws_summary.merge_range(row, 0, row, 5, 'WEIGHT & COST SUMMARY', title_format)
            row += 1

            weight_summary = [
                ['Product Unit Weight:', self.product_id.weight, 'kg'],
                ['Product Total Weight:', self.weight, 'kg'],
                ['', '', ''],
                ['Total Components Weight:', total_component_weight, 'kg'],
                ['Total Materials Weight (from BOMs):', total_material_weight, 'kg'],
                ['', '', ''],
                ['Total Component Cost:', total_component_cost, self.company_id.currency_id.symbol],
                ['Total Material Cost (from BOMs):', total_material_cost, self.company_id.currency_id.symbol],
            ]

            for label, value, unit in weight_summary:
                ws_summary.write(row, 0, label, cell_format)
                if value:
                    if isinstance(value, (int, float)):
                        ws_summary.write(row, 1, value, number_format if 'Weight' in label else currency_format)
                    else:
                        ws_summary.write(row, 1, value, cell_format)
                if unit:
                    ws_summary.write(row, 2, unit, cell_format)
                row += 1

            ws_summary.set_column('A:A', 35)
            ws_summary.set_column('B:B', 20)
            ws_summary.set_column('C:C', 10)

            # ==================== COMPONENTS DETAIL SHEET WITH HEADER ====================
            row = ExcelExportHelper.add_company_header(
                workbook, ws_components, self.env,
                title=f"Components Detail - {self.name}",
                start_row=0
            )

            headers = [
                'Component', 'Quantity', 'UoM', 'Unit Weight (kg)',
                'Total Weight (kg)', 'Cost Price', 'Total Cost', 'Has BOM', 'BOM Code',
                'Specifications'
            ]

            for col, header in enumerate(headers):
                ws_components.write(row, col, header, header_format)

            row += 1
            for comp in self.component_line_ids:
                col = 0
                ws_components.write(row, col, comp.component_id.display_name, cell_format)
                col += 1
                ws_components.write(row, col, comp.quantity, number_format)
                col += 1
                ws_components.write(row, col, comp.uom_id.name, cell_format)
                col += 1
                ws_components.write(row, col, comp.component_id.weight, number_format)
                col += 1
                ws_components.write(row, col, comp.weight * comp.quantity, number_format)
                col += 1
                ws_components.write(row, col, comp.cost_price, currency_format)
                col += 1
                ws_components.write(row, col, comp.total_cost, currency_format)
                col += 1
                ws_components.write(row, col, 'Yes' if comp.bom_id else 'No', cell_format)
                col += 1
                ws_components.write(row, col, comp.bom_id.code if comp.bom_id else '', cell_format)
                col += 1
                specs_text = comp.specifications_display or ''
                ws_components.write(row, col, specs_text, text_wrap_format)
                col += 1
                row += 1

            # Totals
            ws_components.write(row, 0, 'TOTAL', header_format)
            ws_components.write(row, 1, sum(c.quantity for c in self.component_line_ids), number_format)
            ws_components.write(row, 4, total_component_weight, number_format)
            ws_components.write(row, 6, total_component_cost, currency_format)

            ws_components.set_column('A:A', 35)
            ws_components.set_column('B:H', 15)
            ws_components.set_column('I:I', 40)

            # ==================== SPECIFICATIONS SHEET WITH HEADER ====================
            row = ExcelExportHelper.add_company_header(
                workbook, ws_specifications, self.env,
                title=f"Component Specifications - {self.name}",
                start_row=0
            )

            headers = ['Component', 'Specification Type', 'Value', 'Notes']

            for col, header in enumerate(headers):
                ws_specifications.write(row, col, header, header_format)

            row += 1
            for comp in self.component_line_ids:
                if comp.specification_ids:
                    for spec in comp.specification_ids.sorted(lambda s: s.sequence):
                        col = 0
                        ws_specifications.write(row, col, comp.component_id.display_name, cell_format)
                        col += 1
                        ws_specifications.write(row, col, spec.specification_name or '', cell_format)
                        col += 1
                        ws_specifications.write(row, col, spec.value or '', cell_format)
                        col += 1
                        ws_specifications.write(row, col, spec.notes or '', text_wrap_format)
                        col += 1
                        row += 1
                else:
                    col = 0
                    ws_specifications.write(row, col, comp.component_id.display_name, cell_format)
                    col += 1
                    ws_specifications.write(row, col, '-- No Specifications --', cell_format)
                    col += 1
                    ws_specifications.write(row, col, '', cell_format)
                    col += 1
                    ws_specifications.write(row, col, '', cell_format)
                    row += 1

            ws_specifications.set_column('A:A', 35)
            ws_specifications.set_column('B:B', 25)
            ws_specifications.set_column('C:C', 30)
            ws_specifications.set_column('D:D', 40)

            # ==================== MATERIALS BREAKDOWN SHEET WITH HEADER ====================
            row = ExcelExportHelper.add_company_header(
                workbook, ws_materials, self.env,
                title=f"Materials Breakdown - {self.name}",
                start_row=0
            )

            headers = [
                'Component', 'Component Qty', 'BOM Code', 'Material', 'Material Qty per Unit',
                'Total Material Qty', 'Material UoM', 'Material Unit Weight (kg)',
                'Total Material Weight (kg)', 'Unit Cost', 'Total Cost'
            ]

            for col, header in enumerate(headers):
                ws_materials.write(row, col, header, header_format)

            row += 1
            total_mat_weight = 0
            total_mat_cost = 0

            for comp in self.component_line_ids:
                if comp.bom_id:
                    for bom_line in comp.bom_id.bom_line_ids:
                        material = bom_line.product_id
                        material_qty_per_unit = bom_line.product_qty
                        total_material_qty = material_qty_per_unit * comp.quantity
                        material_weight = material.weight * total_material_qty
                        material_cost = material.standard_price * total_material_qty

                        col = 0
                        ws_materials.write(row, col, comp.component_id.display_name, cell_format)
                        col += 1
                        ws_materials.write(row, col, comp.quantity, number_format)
                        col += 1
                        ws_materials.write(row, col, comp.bom_id.code or '', cell_format)
                        col += 1
                        ws_materials.write(row, col, material.display_name, cell_format)
                        col += 1
                        ws_materials.write(row, col, material_qty_per_unit, number_format)
                        col += 1
                        ws_materials.write(row, col, total_material_qty, number_format)
                        col += 1
                        ws_materials.write(row, col, material.uom_id.name, cell_format)
                        col += 1
                        ws_materials.write(row, col, material.weight, number_format)
                        col += 1
                        ws_materials.write(row, col, material_weight, number_format)
                        col += 1
                        ws_materials.write(row, col, material.standard_price, currency_format)
                        col += 1
                        ws_materials.write(row, col, material_cost, currency_format)
                        col += 1

                        total_mat_weight += material_weight
                        total_mat_cost += material_cost
                        row += 1
                else:
                    col = 0
                    ws_materials.write(row, col, comp.component_id.display_name, cell_format)
                    col += 1
                    ws_materials.write(row, col, comp.quantity, number_format)
                    col += 1
                    ws_materials.write(row, col, 'No BOM', cell_format)
                    col += 1
                    ws_materials.write(row, col, comp.component_id.display_name + ' (Direct)', cell_format)
                    col += 1
                    ws_materials.write(row, col, 1, number_format)
                    col += 1
                    ws_materials.write(row, col, comp.quantity, number_format)
                    col += 1
                    ws_materials.write(row, col, comp.uom_id.name, cell_format)
                    col += 1
                    ws_materials.write(row, col, comp.component_id.weight, number_format)
                    col += 1
                    ws_materials.write(row, col, comp.component_id.weight * comp.quantity, number_format)
                    col += 1
                    ws_materials.write(row, col, comp.cost_price, currency_format)
                    col += 1
                    ws_materials.write(row, col, comp.total_cost, currency_format)
                    col += 1

                    total_mat_weight += comp.component_id.weight * comp.quantity
                    total_mat_cost += comp.total_cost
                    row += 1

            # Totals
            ws_materials.write(row, 0, 'TOTAL', header_format)
            ws_materials.write(row, 8, total_mat_weight, number_format)
            ws_materials.write(row, 10, total_mat_cost, currency_format)

            ws_materials.set_column('A:D', 25)
            ws_materials.set_column('E:K', 15)

            workbook.close()
            output.seek(0)

            file_data = base64.b64encode(output.read())
            filename = 'Product_Pricing_%s.xlsx' % self.name.replace('/', '_')

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

        except ImportError:
            raise UserError(_('Please install xlsxwriter library: pip install xlsxwriter'))
        except Exception as e:
            raise UserError(_('Error creating Excel file: %s') % str(e))


class ProjectProductComponent(models.Model):
    _name = 'project.product.component'
    _description = 'Project Product Component'
    _order = 'sequence, id'

    sequence = fields.Integer(string='Sequence', default=10)
    pricing_id = fields.Many2one(
        'project.product.pricing',
        string='Pricing',
        required=True,
        ondelete='cascade'
    )
    component_id = fields.Many2one(
        'product.product',
        string='Component Product',
        required=True,
        domain=[('type', 'in', ['product', 'consu'])]
    )
    quantity = fields.Float(
        string='Quantity',
        required=True,
        default=1.0,
        digits='Product Unit of Measure'
    )
    weight = fields.Float(
        string='Weight',
        digits='Stock Weight'
    )
    cost_price = fields.Float(
        string='Cost Price',
        required=True,
        digits='Product Price'
    )
    uom_id = fields.Many2one(
        'uom.uom',
        string='Unit of Measure',
        related='component_id.uom_id',
        readonly=True
    )
    total_cost = fields.Float(
        string='Total Cost',
        compute='_compute_total_cost',
        store=True,
        digits='Product Price'
    )
    bom_id = fields.Many2one(
        'mrp.bom',
        string='Bill of Materials'
    )
    specification_ids = fields.One2many(
        'component.specification.value',
        'pricing_component_id',
        string='Specifications'
    )
    spec_count = fields.Integer(
        string='Specifications',
        compute='_compute_spec_count'
    )

    specifications_display = fields.Text(
        string='Specifications',
        compute='_compute_specifications_display',
        store=False,
        help='Display of all specifications'
    )

    @api.depends('specification_ids')
    def _compute_spec_count(self):
        for record in self:
            record.spec_count = len(record.specification_ids)

    @api.depends('specification_ids', 'specification_ids.specification_name', 'specification_ids.value')
    def _compute_specifications_display(self):
        """Compute display text for specifications"""
        for record in self:
            if record.specification_ids:
                specs = []
                for spec in record.specification_ids.sorted(lambda s: s.sequence):
                    specs.append(f"{spec.specification_name}: {spec.value}")
                record.specifications_display = '\n'.join(specs)
            else:
                record.specifications_display = ''

    @api.depends('quantity', 'cost_price')
    def _compute_total_cost(self):
        for line in self:
            line.total_cost = line.quantity * line.cost_price

    @api.onchange('component_id')
    def _onchange_component_id(self):
        if self.component_id:
            self.cost_price = self.component_id.standard_price
            self.weight = self.component_id.weight
            bom = self.env['mrp.bom'].search([
                ('product_id', '=', self.component_id.id)
            ], limit=1)
            if bom:
                self.bom_id = bom.id

    def action_view_bom(self):
        self.ensure_one()
        if not self.bom_id:
            return {
                'type': 'ir.actions.act_window',
                'name': 'Create Bill of Materials',
                'res_model': 'mrp.bom',
                'view_mode': 'form',
                'context': {
                    'default_product_id': self.component_id.id,
                    'default_product_tmpl_id': self.component_id.product_tmpl_id.id,
                },
                'target': 'new',
            }
        return {
            'type': 'ir.actions.act_window',
            'name': 'Bill of Materials',
            'res_model': 'mrp.bom',
            'view_mode': 'form',
            'res_id': self.bom_id.id,
            'target': 'current',
        }

    def action_create_bom(self):
        self.ensure_one()
        bom = self.env['mrp.bom'].create({
            'product_id': self.component_id.id,
            'product_tmpl_id': self.component_id.product_tmpl_id.id,
            'product_qty': 1.0,
            'type': 'normal',
        })
        self.bom_id = bom.id
        return self.action_view_bom()

    def action_component_specifications(self):
        """Open specifications wizard"""
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _('Component Specifications: %s') % self.component_id.name,
            'res_model': 'component.specification.wizard',
            'view_mode': 'form',
            'target': 'new',
            'context': {
                'default_component_id': self.component_id.id,
                'source_model': 'project.product.component',
                'source_id': self.id,
                'component_id': self.component_id.id,
            }
        }