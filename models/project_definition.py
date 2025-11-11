# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import ValidationError, UserError
import logging
import base64
import io

_logger = logging.getLogger(__name__)


class ProjectDefinition(models.Model):
    _name = 'project.definition'
    _description = 'Project Definition'
    _inherit = ['mail.thread', 'mail.activity.mixin']
    _order = 'create_date desc'

    name = fields.Char(
        string='Project Code',
        required=True,
        copy=False,
        readonly=True,
        default=lambda self: _('New'),
        tracking=True
    )
    project_name = fields.Char(
        string='Project Name',
        required=True,
        tracking=True
    )
    partner_id = fields.Many2one(
        'res.partner',
        string='Customer',
        required=True,
        domain=[('customer_rank', '>', 0)],
        tracking=True
    )
    start_date = fields.Date(
        string='Project Start Date',
        required=True,
        default=fields.Date.context_today,
        tracking=True
    )
    end_date = fields.Date(
        string='Expected End Date',
        required=True,
        tracking=True
    )
    state = fields.Selection([
        ('draft', 'Draft'),
        ('pricing', 'Pricing'),
        ('planning', 'Planning'),
        ('processing', 'Processing'),
        ('done', 'Done'),
        ('cancelled', 'Cancelled'),
    ], string='Status', default='draft', tracking=True)

    product_line_ids = fields.One2many(
        'project.product.line',
        'project_id',
        string='Project Products'
    )

    total_cost = fields.Float(
        string='Total Cost',
        compute='_compute_totals',
        store=True
    )
    total_sale = fields.Float(
        string='Total Sale',
        compute='_compute_totals',
        store=True
    )
    total_profit = fields.Float(
        string='Total Profit',
        compute='_compute_totals',
        store=True
    )

    company_id = fields.Many2one(
        'res.company',
        string='Company',
        default=lambda self: self.env.company
    )

    notes = fields.Text(string='Notes')

    # Related records counts
    pricing_count = fields.Integer(
        string='Pricings',
        compute='_compute_related_counts'
    )
    planning_count = fields.Integer(
        string='Plannings',
        compute='_compute_related_counts'
    )
    sales_order_count = fields.Integer(
        string='Sales Orders',
        compute='_compute_related_counts'
    )
    execution_count = fields.Integer(
        string='Work Order Executions',
        compute='_compute_related_counts'
    )
    estimation_count = fields.Integer(
        string='Cost Estimations',
        compute='_compute_related_counts'
    )

    # Auto-update control
    auto_update_state = fields.Boolean(
        string='Auto-Update State',
        default=True,
        tracking=True,
        help='Automatically update project state based on activities'
    )

    @api.model
    def create(self, vals):
        # Auto-generate sequence
        if vals.get('name', _('New')) == _('New'):
            vals['name'] = self.env['ir.sequence'].next_by_code('project.definition') or _('New')

        # Ensure partner is marked as customer
        if vals.get('partner_id'):
            partner = self.env['res.partner'].browse(vals['partner_id'])
            if partner.customer_rank == 0:
                _logger.info('Auto-marking partner %s as customer', partner.name)
                partner.write({'customer_rank': 1})

        return super(ProjectDefinition, self).create(vals)

    def write(self, vals):
        # Ensure partner is marked as customer when changed
        if vals.get('partner_id'):
            partner = self.env['res.partner'].browse(vals['partner_id'])
            if partner.customer_rank == 0:
                _logger.info('Auto-marking partner %s as customer', partner.name)
                partner.write({'customer_rank': 1})

        return super(ProjectDefinition, self).write(vals)

    @api.depends('product_line_ids.cost_price', 'product_line_ids.sale_price', 'product_line_ids.quantity')
    def _compute_totals(self):
        for record in self:
            total_cost = sum(line.cost_price * line.quantity for line in record.product_line_ids)
            total_sale = sum(line.sale_price * line.quantity for line in record.product_line_ids)
            record.total_cost = total_cost
            record.total_sale = total_sale
            record.total_profit = total_sale - total_cost

    def _compute_related_counts(self):
        for record in self:
            record.pricing_count = self.env['project.product.pricing'].search_count([
                ('project_id', '=', record.id)
            ])
            record.planning_count = self.env['material.production.planning'].search_count([
                ('project_id', '=', record.id)
            ])

            # Count sales orders by origin
            sales_orders = self.env['sale.order'].search([
                ('client_order_ref', '=', record.name)
            ])
            record.sales_order_count = len(sales_orders)

            record.execution_count = self.env['work.order.execution'].search_count([
                ('project_id', '=', record.id)
            ])

            # Count cost estimations
            record.estimation_count = self.env['project.cost.estimation'].search_count([
                ('project_id', '=', record.id)
            ])

    @api.constrains('start_date', 'end_date')
    def _check_dates(self):
        for record in self:
            if record.end_date and record.start_date and record.end_date < record.start_date:
                raise ValidationError(_('End date cannot be before start date!'))

    # ==================== PUBLIC METHOD FOR TRIGGERING UPDATES ====================

    def update_project_state(self):
        """Public method to trigger state update from related models"""
        for project in self:
            project._auto_update_state()

    # ==================== AUTOMATIC STATE UPDATES ====================

    def _auto_update_state(self):
        """Automatically update project state based on actual activities"""
        self.ensure_one()

        if not self.auto_update_state:
            return

        # Skip if already done or cancelled
        if self.state in ('done', 'cancelled'):
            return

        # Get related records
        pricings = self.env['project.product.pricing'].search([
            ('project_id', '=', self.id)
        ])
        plannings = self.env['material.production.planning'].search([
            ('project_id', '=', self.id)
        ])
        executions = self.env['work.order.execution'].search([
            ('project_id', '=', self.id)
        ])

        # Get production orders
        productions = self.env['mrp.production'].search([
            ('origin', 'ilike', self.name)
        ])

        # Priority 1: Check if all work is done
        if executions and all(exe.state == 'done' for exe in executions):
            if plannings and all(plan.state == 'done' for plan in plannings):
                if productions and all(prod.state == 'done' for prod in productions):
                    self._move_to_done()
                    return

        # Priority 2: Check if work orders are in progress
        if executions and any(exe.state in ('loaded', 'in_progress') for exe in executions):
            self._move_to_processing()
            return

        # Check if any productions are in progress
        if productions and any(prod.state in ('confirmed', 'progress', 'to_close') for prod in productions):
            self._move_to_processing()
            return

        # Priority 3: Check if planning is created and work orders exist
        if plannings and any(plan.state in ('work_orders_created', 'done') for plan in plannings):
            if self.state not in ('planning', 'processing', 'done'):
                self._move_to_planning()
            return

        # Priority 4: Check if planning exists (even in earlier states)
        if plannings and any(plan.state != 'draft' for plan in plannings):
            if self.state not in ('planning', 'processing', 'done'):
                self._move_to_planning()
            return

        # Priority 5: Check if pricing is confirmed/approved
        if pricings and any(prc.state in ('confirmed', 'approved') for prc in pricings):
            if self.state == 'draft':
                self._move_to_pricing()
            return

        # Priority 6: Check if pricing exists (even in draft)
        if pricings and any(prc.state != 'draft' for prc in pricings):
            if self.state == 'draft':
                self._move_to_pricing()
            return

    def _move_to_pricing(self):
        """Move project to pricing state"""
        if self.state == 'draft':
            self.write({'state': 'pricing'})
            self.message_post(
                body=_('🔄 Project automatically moved to Pricing stage (Pricing created)'),
                subtype_xmlid='mail.mt_note'
            )
            _logger.info('Project %s auto-moved to Pricing', self.name)

    def _move_to_planning(self):
        """Move project to planning state"""
        if self.state in ('draft', 'pricing'):
            self.write({'state': 'planning'})
            self.message_post(
                body=_('🔄 Project automatically moved to Planning stage (Material Planning created)'),
                subtype_xmlid='mail.mt_note'
            )
            _logger.info('Project %s auto-moved to Planning', self.name)

    def _move_to_processing(self):
        """Move project to processing state"""
        if self.state in ('draft', 'pricing', 'planning'):
            self.write({'state': 'processing'})
            self.message_post(
                body=_('🔄 Project automatically moved to Processing stage (Work Orders started)'),
                subtype_xmlid='mail.mt_note'
            )
            _logger.info('Project %s auto-moved to Processing', self.name)

    def _move_to_done(self):
        """Move project to done state"""
        if self.state != 'done':
            self.write({'state': 'done'})
            self.message_post(
                body=_('✅ Project automatically marked as Done (All work orders completed)'),
                subtype_xmlid='mail.mt_comment'
            )
            _logger.info('Project %s auto-moved to Done', self.name)

    # ==================== MANUAL STATE CHANGES ====================

    def action_pricing(self):
        """Move to Pricing stage"""
        self.write({'state': 'pricing'})

    def action_planning(self):
        """Move to Planning stage"""
        self.write({'state': 'planning'})

    def action_processing(self):
        """Move to Processing stage"""
        self.write({'state': 'processing'})

    def action_done(self):
        """Mark as Done"""
        self.write({'state': 'done'})

    def action_cancel(self):
        self.write({'state': 'cancelled'})

    def action_draft(self):
        self.write({'state': 'draft'})

    # ==================== VIEW ACTIONS ====================

    def action_view_pricings(self):
        """View related pricings"""
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _('Product Pricings'),
            'res_model': 'project.product.pricing',
            'view_mode': 'tree,form',
            'domain': [('project_id', '=', self.id)],
            'context': {
                'default_project_id': self.id,
                'default_partner_id': self.partner_id.id,
            },
        }

    def action_view_plannings(self):
        """View related material plannings"""
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _('Material Plannings'),
            'res_model': 'material.production.planning',
            'view_mode': 'tree,form',
            'domain': [('project_id', '=', self.id)],
            'context': {
                'default_project_id': self.id,
            },
        }

    def action_view_sales_orders(self):
        """View related sales orders"""
        self.ensure_one()
        sales_orders = self.env['sale.order'].search([
            ('client_order_ref', '=', self.name)
        ])
        return {
            'type': 'ir.actions.act_window',
            'name': _('Sales Orders'),
            'res_model': 'sale.order',
            'view_mode': 'tree,form',
            'domain': [('id', 'in', sales_orders.ids)],
            'context': {
                'default_partner_id': self.partner_id.id,
            },
        }

    def action_view_executions(self):
        """View work order executions"""
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _('Work Order Executions'),
            'res_model': 'work.order.execution',
            'view_mode': 'tree,form',
            'domain': [('project_id', '=', self.id)],
            'context': {
                'default_project_id': self.id,
            },
        }

    # ==================== COST ESTIMATION ACTIONS ====================

    def action_create_cost_estimation(self):
        """Create new cost estimation"""
        self.ensure_one()

        if not self.product_line_ids:
            raise UserError(_('Please add products to the project first!'))

        # Create estimation record
        estimation = self.env['project.cost.estimation'].create({
            'project_id': self.id,
        })

        # Generate Excel file
        estimation.action_generate_estimation_excel()

        # Open estimation form
        return {
            'type': 'ir.actions.act_window',
            'name': _('Cost Estimation'),
            'res_model': 'project.cost.estimation',
            'res_id': estimation.id,
            'view_mode': 'form',
            'target': 'current',
        }

    def action_view_estimations(self):
        """View all cost estimations"""
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _('Cost Estimations'),
            'res_model': 'project.cost.estimation',
            'view_mode': 'tree,form',
            'domain': [('project_id', '=', self.id)],
            'context': {
                'default_project_id': self.id,
            },
        }

    # ==================== EXCEL EXPORT ====================

    def action_export_project_excel(self):
        """Export project details to Excel with company header"""
        self.ensure_one()

        try:
            import xlsxwriter
        except ImportError:
            raise UserError(_('Please install xlsxwriter library: pip install xlsxwriter'))

        try:
            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})

            # Create worksheets
            ws_summary = workbook.add_worksheet('Project Summary')
            ws_products = workbook.add_worksheet('Project Products')
            ws_status = workbook.add_worksheet('Project Status')

            # Formats
            title_format = workbook.add_format({
                'bold': True,
                'font_size': 18,
                'font_color': '#1F4E78',
                'align': 'left',
            })

            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4CAF50',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
            })

            section_header_format = workbook.add_format({
                'bold': True,
                'font_size': 12,
                'bg_color': '#2E86C1',
                'font_color': 'white',
                'border': 1,
            })

            cell_format = workbook.add_format({
                'border': 1,
                'align': 'left',
                'valign': 'vcenter',
            })

            number_format = workbook.add_format({
                'border': 1,
                'align': 'right',
                'num_format': '#,##0.00',
            })

            currency_format = workbook.add_format({
                'border': 1,
                'align': 'right',
                'num_format': '#,##0.00',
            })

            date_format = workbook.add_format({
                'border': 1,
                'align': 'center',
                'num_format': 'yyyy-mm-dd',
            })

            # ==================== PROJECT SUMMARY SHEET ====================
            company = self.env.company
            row = 0

            # Company logo
            if company.logo:
                try:
                    logo_data = base64.b64decode(company.logo)
                    image_data = io.BytesIO(logo_data)
                    ws_summary.insert_image(row, 0, 'logo.png', {
                        'x_scale': 0.5,
                        'y_scale': 0.5,
                        'x_offset': 10,
                        'y_offset': 10,
                        'image_data': image_data,
                    })
                except Exception as e:
                    _logger.warning('Could not insert company logo: %s', str(e))

            # Company info
            ws_summary.merge_range(row, 2, row, 5, company.name or 'Company Name', title_format)
            row += 1

            if company.street or company.city:
                address = ', '.join(filter(None, [company.street, company.city, company.country_id.name]))
                ws_summary.merge_range(row, 2, row, 5, address, cell_format)
                row += 1

            if company.phone or company.email:
                contact = ' | '.join(filter(None, [f"Tel: {company.phone}" if company.phone else None,
                                                   f"Email: {company.email}" if company.email else None]))
                ws_summary.merge_range(row, 2, row, 5, contact, cell_format)
                row += 1

            row += 1
            ws_summary.merge_range(row, 0, row, 5, f'PROJECT REPORT - {self.name}', title_format)
            row += 1

            from datetime import datetime
            export_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            ws_summary.merge_range(row, 0, row, 5, f'Export Date: {export_date}', cell_format)
            row += 2

            # Project Information
            ws_summary.merge_range(row, 0, row, 5, 'PROJECT INFORMATION', section_header_format)
            row += 1

            project_info = [
                ['Project Code:', self.name],
                ['Project Name:', self.project_name],
                ['Customer:', self.partner_id.name],
                ['Status:', dict(self._fields['state'].selection).get(self.state)],
                ['Start Date:', self.start_date.strftime('%Y-%m-%d') if self.start_date else ''],
                ['End Date:', self.end_date.strftime('%Y-%m-%d') if self.end_date else ''],
                ['Auto-Update State:', 'Enabled' if self.auto_update_state else 'Disabled'],
            ]

            for label, value in project_info:
                ws_summary.write(row, 0, label, cell_format)
                ws_summary.write(row, 1, value, cell_format)
                row += 1

            row += 1

            # Financial Summary
            ws_summary.merge_range(row, 0, row, 5, 'FINANCIAL SUMMARY', section_header_format)
            row += 1

            ws_summary.write(row, 0, 'Total Cost:', cell_format)
            ws_summary.write(row, 1, self.total_cost, currency_format)
            ws_summary.write(row, 2, self.company_id.currency_id.symbol, cell_format)
            row += 1

            ws_summary.write(row, 0, 'Total Sale:', cell_format)
            ws_summary.write(row, 1, self.total_sale, currency_format)
            ws_summary.write(row, 2, self.company_id.currency_id.symbol, cell_format)
            row += 1

            ws_summary.write(row, 0, 'Total Profit:', cell_format)
            ws_summary.write(row, 1, self.total_profit, currency_format)
            ws_summary.write(row, 2, self.company_id.currency_id.symbol, cell_format)
            row += 1

            profit_margin = (self.total_profit / self.total_sale * 100) if self.total_sale > 0 else 0
            ws_summary.write(row, 0, 'Profit Margin:', cell_format)
            ws_summary.write(row, 1, profit_margin, number_format)
            ws_summary.write(row, 2, '%', cell_format)
            row += 1

            ws_summary.set_column('A:A', 25)
            ws_summary.set_column('B:B', 30)
            ws_summary.set_column('C:E', 15)

            # ==================== PRODUCTS SHEET ====================
            row = 0
            ws_products.merge_range(row, 0, row, 8, f'PROJECT PRODUCTS - {self.name}', title_format)
            row += 2

            headers = [
                'Product', 'Quantity', 'UoM', 'Weight (kg)',
                'Cost Price', 'Sale Price', 'Total Cost', 'Total Sale', 'Profit'
            ]

            for col, header in enumerate(headers):
                ws_products.write(row, col, header, header_format)

            row += 1

            for product_line in self.product_line_ids:
                col = 0
                ws_products.write(row, col, product_line.product_id.display_name, cell_format)
                col += 1
                ws_products.write(row, col, product_line.quantity, number_format)
                col += 1
                ws_products.write(row, col, product_line.uom_id.name, cell_format)
                col += 1
                ws_products.write(row, col, product_line.weight, number_format)
                col += 1
                ws_products.write(row, col, product_line.cost_price, currency_format)
                col += 1
                ws_products.write(row, col, product_line.sale_price, currency_format)
                col += 1
                ws_products.write(row, col, product_line.total_cost, currency_format)
                col += 1
                ws_products.write(row, col, product_line.total_sale, currency_format)
                col += 1
                ws_products.write(row, col, product_line.profit, currency_format)
                row += 1

            # Totals
            ws_products.write(row, 0, 'TOTAL', header_format)
            ws_products.write(row, 6, self.total_cost, currency_format)
            ws_products.write(row, 7, self.total_sale, currency_format)
            ws_products.write(row, 8, self.total_profit, currency_format)

            ws_products.set_column('A:A', 35)
            ws_products.set_column('B:H', 15)

            # ==================== STATUS SHEET ====================
            row = 0
            ws_status.merge_range(row, 0, row, 5, f'PROJECT STATUS DETAILS - {self.name}', title_format)
            row += 2

            # Pricings
            ws_status.merge_range(row, 0, row, 5, 'PRODUCT PRICINGS', section_header_format)
            row += 1

            pricings = self.env['project.product.pricing'].search([
                ('project_id', '=', self.id)
            ])

            if pricings:
                headers = ['Pricing Code', 'Product', 'Version', 'Status', 'Date', 'Total Cost']
                for col, header in enumerate(headers):
                    ws_status.write(row, col, header, header_format)
                row += 1

                for pricing in pricings:
                    ws_status.write(row, 0, pricing.name, cell_format)
                    ws_status.write(row, 1, pricing.product_id.display_name, cell_format)
                    ws_status.write(row, 2, pricing.version, number_format)
                    ws_status.write(row, 3, dict(pricing._fields['state'].selection).get(pricing.state), cell_format)
                    ws_status.write(row, 4, pricing.pricing_date.strftime('%Y-%m-%d') if pricing.pricing_date else '',
                                    date_format)
                    ws_status.write(row, 5, pricing.total_component_cost, currency_format)
                    row += 1
            else:
                ws_status.merge_range(row, 0, row, 5, 'No pricings created yet', cell_format)
                row += 1

            row += 1

            # Material Plannings
            ws_status.merge_range(row, 0, row, 5, 'MATERIAL PLANNINGS', section_header_format)
            row += 1

            plannings = self.env['material.production.planning'].search([
                ('project_id', '=', self.id)
            ])

            if plannings:
                headers = ['Planning Reference', 'Product', 'Quantity', 'Status', 'Production Orders']
                for col, header in enumerate(headers):
                    ws_status.write(row, col, header, header_format)
                row += 1

                for planning in plannings:
                    ws_status.write(row, 0, planning.name, cell_format)
                    ws_status.write(row, 1, planning.product_id.display_name, cell_format)
                    ws_status.write(row, 2, planning.quantity, number_format)
                    ws_status.write(row, 3, dict(planning._fields['state'].selection).get(planning.state), cell_format)
                    ws_status.write(row, 4, planning.production_count, number_format)
                    row += 1
            else:
                ws_status.merge_range(row, 0, row, 5, 'No material plannings created yet', cell_format)
                row += 1

            row += 1

            # Sales Orders
            ws_status.merge_range(row, 0, row, 5, 'SALES ORDERS', section_header_format)
            row += 1

            sales_orders = self.env['sale.order'].search([
                ('client_order_ref', '=', self.name)
            ])

            if sales_orders:
                headers = ['Order Reference', 'Date', 'Status', 'Total Amount', 'Currency']
                for col, header in enumerate(headers):
                    ws_status.write(row, col, header, header_format)
                row += 1

                for so in sales_orders:
                    ws_status.write(row, 0, so.name, cell_format)
                    ws_status.write(row, 1, so.date_order.strftime('%Y-%m-%d') if so.date_order else '', date_format)
                    ws_status.write(row, 2, dict(so._fields['state'].selection).get(so.state), cell_format)
                    ws_status.write(row, 3, so.amount_total, currency_format)
                    ws_status.write(row, 4, so.currency_id.symbol, cell_format)
                    row += 1
            else:
                ws_status.merge_range(row, 0, row, 5, 'No sales orders created yet', cell_format)
                row += 1

            row += 1

            # Work Order Executions
            ws_status.merge_range(row, 0, row, 5, 'WORK ORDER EXECUTIONS', section_header_format)
            row += 1

            executions = self.env['work.order.execution'].search([
                ('project_id', '=', self.id)
            ])

            if executions:
                headers = ['Execution Reference', 'Product', 'Status', 'Total Components', 'Completed']
                for col, header in enumerate(headers):
                    ws_status.write(row, col, header, header_format)
                row += 1

                for exe in executions:
                    ws_status.write(row, 0, exe.name, cell_format)
                    ws_status.write(row, 1, exe.product_id.display_name, cell_format)
                    ws_status.write(row, 2, dict(exe._fields['state'].selection).get(exe.state), cell_format)
                    ws_status.write(row, 3, exe.total_components, number_format)
                    ws_status.write(row, 4, exe.completed_components, number_format)
                    row += 1
            else:
                ws_status.merge_range(row, 0, row, 5, 'No work order executions created yet', cell_format)
                row += 1

            row += 1

            # Cost Estimations
            ws_status.merge_range(row, 0, row, 5, 'COST ESTIMATIONS', section_header_format)
            row += 1

            estimations = self.env['project.cost.estimation'].search([
                ('project_id', '=', self.id)
            ])

            if estimations:
                headers = ['Estimation Ref', 'Date', 'Status', 'Last Update']
                for col, header in enumerate(headers):
                    ws_status.write(row, col, header, header_format)
                row += 1

                for est in estimations:
                    ws_status.write(row, 0, est.name, cell_format)
                    ws_status.write(row, 1,
                                    est.estimation_date.strftime('%Y-%m-%d %H:%M') if est.estimation_date else '',
                                    cell_format)
                    ws_status.write(row, 2, dict(est._fields['state'].selection).get(est.state), cell_format)
                    ws_status.write(row, 3,
                                    est.last_update_date.strftime('%Y-%m-%d %H:%M') if est.last_update_date else '',
                                    cell_format)
                    row += 1
            else:
                ws_status.merge_range(row, 0, row, 5, 'No cost estimations created yet', cell_format)
                row += 1

            ws_status.set_column('A:A', 25)
            ws_status.set_column('B:E', 20)

            workbook.close()
            output.seek(0)

            file_data = base64.b64encode(output.read())
            filename = f'Project_{self.name.replace("/", "_")}.xlsx'

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

        except Exception as e:
            raise UserError(_('Error creating Excel file: %s') % str(e))


class ProjectProductLine(models.Model):
    _name = 'project.product.line'
    _description = 'Project Product Line'
    _order = 'sequence, id'

    sequence = fields.Integer(string='Sequence', default=10)
    project_id = fields.Many2one(
        'project.definition',
        string='Project',
        required=True,
        ondelete='cascade'
    )
    product_id = fields.Many2one(
        'product.product',
        string='Product',
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
    sale_price = fields.Float(
        string='Sale Price',
        required=True,
        digits='Product Price'
    )
    uom_id = fields.Many2one(
        'uom.uom',
        string='Unit of Measure',
        related='product_id.uom_id',
        readonly=True
    )
    total_cost = fields.Float(
        string='Total Cost',
        compute='_compute_total',
        store=True,
        digits='Product Price'
    )
    total_sale = fields.Float(
        string='Total Sale',
        compute='_compute_total',
        store=True,
        digits='Product Price'
    )
    profit = fields.Float(
        string='Profit',
        compute='_compute_total',
        store=True,
        digits='Product Price'
    )

    @api.depends('quantity', 'cost_price', 'sale_price')
    def _compute_total(self):
        for line in self:
            line.total_cost = line.quantity * line.cost_price
            line.total_sale = line.quantity * line.sale_price
            line.profit = line.total_sale - line.total_cost

    @api.onchange('product_id')
    def _onchange_product_id(self):
        if self.product_id:
            self.cost_price = self.product_id.standard_price
            self.sale_price = self.product_id.list_price
            self.weight = self.product_id.weight