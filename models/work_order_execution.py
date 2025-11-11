# -*- coding: utf-8 -*-

from odoo import models, fields, api, tools, _
from odoo.exceptions import UserError, ValidationError
import logging
import base64
import io

_logger = logging.getLogger(__name__)


class ExcelExportHelper:
    """Helper class for Excel exports with company header - FIXED VERSION"""

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

                # FIXED: Insert logo directly from bytes (no temp file needed)
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


class WorkOrderExecution(models.Model):
    _name = 'work.order.execution'
    _description = 'Work Order Execution'
    _inherit = ['mail.thread', 'mail.activity.mixin']
    _order = 'create_date desc'

    name = fields.Char(
        string='Execution Reference',
        required=True,
        copy=False,
        readonly=True,
        default=lambda self: _('New'),
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
    state = fields.Selection([
        ('draft', 'Draft'),
        ('loaded', 'Work Orders Loaded'),
        ('in_progress', 'In Progress'),
        ('done', 'Done'),
        ('cancelled', 'Cancelled'),
    ], string='Status', default='draft', tracking=True)

    work_order_line_ids = fields.One2many(
        'work.order.execution.line',
        'execution_id',
        string='Work Order Lines'
    )

    company_id = fields.Many2one(
        'res.company',
        string='Company',
        default=lambda self: self.env.company
    )

    notes = fields.Text(string='Notes')

    total_components = fields.Integer(
        string='Total Components',
        compute='_compute_totals',
        store=True
    )
    completed_components = fields.Integer(
        string='Completed Components',
        compute='_compute_totals',
        store=True
    )
    in_progress_components = fields.Integer(
        string='In Progress',
        compute='_compute_totals',
        store=True
    )

    material_issue_count = fields.Integer(
        string='Material Issues',
        compute='_compute_material_issue_count'
    )
    operation_actual_count = fields.Integer(
        string='Operations Actual',
        compute='_compute_operation_actual_count'
    )

    @api.model
    def create(self, vals):
        if vals.get('name', _('New')) == _('New'):
            vals['name'] = self.env['ir.sequence'].next_by_code('work.order.execution') or _('New')
        return super(WorkOrderExecution, self).create(vals)

    @api.depends('work_order_line_ids.production_state')
    def _compute_totals(self):
        for record in self:
            record.total_components = len(record.work_order_line_ids)
            record.completed_components = len(record.work_order_line_ids.filtered(
                lambda l: l.production_state == 'done'
            ))
            record.in_progress_components = len(record.work_order_line_ids.filtered(
                lambda l: l.production_state in ('confirmed', 'progress', 'to_close')
            ))

    def _compute_material_issue_count(self):
        """Count material moves (issues) for selected work orders"""
        for record in self:
            selected_lines = record.work_order_line_ids.filtered(lambda l: l.selected)
            if selected_lines:
                production_ids = selected_lines.mapped('production_id').ids

                moves = self.env['stock.move'].search([
                    '|',
                    ('raw_material_production_id', 'in', production_ids),
                    '&',
                    ('production_id', 'in', production_ids),
                    ('location_dest_id.usage', '=', 'production')
                ])

                if not moves:
                    move_lines = self.env['stock.move.line'].search([
                        ('production_id', 'in', production_ids),
                        ('location_dest_id.usage', '=', 'production')
                    ])
                    moves = move_lines.mapped('move_id')

                record.material_issue_count = len(moves)
            else:
                record.material_issue_count = 0

    def _compute_operation_actual_count(self):
        """Count workorders for selected work orders"""
        for record in self:
            selected_lines = record.work_order_line_ids.filtered(lambda l: l.selected)
            if selected_lines:
                production_ids = selected_lines.mapped('production_id').ids
                workorders = self.env['mrp.workorder'].search([
                    ('production_id', 'in', production_ids),
                ])
                record.operation_actual_count = len(workorders)
            else:
                record.operation_actual_count = 0

    @api.onchange('project_id')
    def _onchange_project_id(self):
        self.product_id = False
        if self.project_id:
            product_ids = self.project_id.product_line_ids.mapped('product_id').ids
            return {'domain': {'product_id': [('id', 'in', product_ids)]}}
        return {'domain': {'product_id': []}}

    def action_load_work_orders(self):
        self.ensure_one()

        if not self.product_id or not self.project_id:
            raise UserError(_('Please select Project and Product first!'))

        self.work_order_line_ids.unlink()

        planning = self.env['material.production.planning'].search([
            ('project_id', '=', self.project_id.id),
            ('product_id', '=', self.product_id.id),
            ('state', 'in', ['work_orders_created', 'done'])
        ], limit=1, order='create_date desc')

        if not planning:
            draft_planning = self.env['material.production.planning'].search([
                ('project_id', '=', self.project_id.id),
                ('product_id', '=', self.product_id.id),
            ], limit=1, order='create_date desc')

            if draft_planning:
                raise UserError(_(
                    'Material Planning exists but no work orders created yet!\n\n'
                    'Planning: %s\n'
                    'State: %s\n\n'
                    'Please go to Material Planning and create work orders first.'
                ) % (draft_planning.name, dict(draft_planning._fields['state'].selection).get(draft_planning.state)))
            else:
                raise UserError(_(
                    'No Material Planning found for this project and product!\n\n'
                    'Please create a Material Planning first.'
                ))

        if not planning.production_order_ids:
            raise UserError(_('Material Planning exists but no production orders found!'))

        productions = planning.production_order_ids

        total_operations = 0
        for production in productions:
            if production.state == 'draft':
                production.action_confirm()

            if not production.workorder_ids and production.state in ('confirmed', 'progress'):
                try:
                    production._create_workorder()
                    _logger.info('Created workorders for production %s', production.name)
                except Exception as e:
                    _logger.warning('Could not create workorders for %s: %s', production.name, str(e))

            line = self.env['work.order.execution.line'].create({
                'execution_id': self.id,
                'component_id': production.product_id.id,
                'quantity': production.product_qty,
                'weight': production.product_id.weight * production.product_qty,
                'production_id': production.id,
            })

            ops_created = self._load_operations_for_line(line)
            total_operations += ops_created

        self.state = 'loaded'

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Success'),
                'message': _('%s work orders loaded with %s operations!') % (len(productions), total_operations),
                'type': 'success',
                'sticky': False,
            }
        }

    @api.model
    def create(self, vals):
        """Trigger project state update when execution is created"""
        if vals.get('name', _('New')) == _('New'):
            vals['name'] = self.env['ir.sequence'].next_by_code('work.order.execution') or _('New')

        execution = super(WorkOrderExecution, self).create(vals)

        # Update project state
        if execution.project_id and execution.project_id.auto_update_state:
            execution.project_id.update_project_state()

        return execution

    def write(self, vals):
        """Trigger project state update when execution state changes"""
        result = super(WorkOrderExecution, self).write(vals)

        # Update project state if state changed
        if 'state' in vals:
            for execution in self:
                if execution.project_id and execution.project_id.auto_update_state:
                    execution.project_id.update_project_state()

        return result

    def action_load_work_orders(self):
        """Load work orders and update project state"""
        self.ensure_one()

        if not self.product_id or not self.project_id:
            raise UserError(_('Please select Project and Product first!'))

        self.work_order_line_ids.unlink()

        planning = self.env['material.production.planning'].search([
            ('project_id', '=', self.project_id.id),
            ('product_id', '=', self.product_id.id),
            ('state', 'in', ['work_orders_created', 'done'])
        ], limit=1, order='create_date desc')

        if not planning:
            draft_planning = self.env['material.production.planning'].search([
                ('project_id', '=', self.project_id.id),
                ('product_id', '=', self.product_id.id),
            ], limit=1, order='create_date desc')

            if draft_planning:
                raise UserError(_(
                    'Material Planning exists but no work orders created yet!\n\n'
                    'Planning: %s\n'
                    'State: %s\n\n'
                    'Please go to Material Planning and create work orders first.'
                ) % (draft_planning.name, dict(draft_planning._fields['state'].selection).get(draft_planning.state)))
            else:
                raise UserError(_(
                    'No Material Planning found for this project and product!\n\n'
                    'Please create a Material Planning first.'
                ))

        if not planning.production_order_ids:
            raise UserError(_('Material Planning exists but no production orders found!'))

        productions = planning.production_order_ids

        total_operations = 0
        for production in productions:
            if production.state == 'draft':
                production.action_confirm()

            if not production.workorder_ids and production.state in ('confirmed', 'progress'):
                try:
                    production._create_workorder()
                    _logger.info('Created workorders for production %s', production.name)
                except Exception as e:
                    _logger.warning('Could not create workorders for %s: %s', production.name, str(e))

            line = self.env['work.order.execution.line'].create({
                'execution_id': self.id,
                'component_id': production.product_id.id,
                'quantity': production.product_qty,
                'weight': production.product_id.weight * production.product_qty,
                'production_id': production.id,
            })

            ops_created = self._load_operations_for_line(line)
            total_operations += ops_created

        self.state = 'loaded'

        # Trigger project state update
        if self.project_id and self.project_id.auto_update_state:
            self.project_id.update_project_state()

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Success'),
                'message': _('%s work orders loaded with %s operations!') % (len(productions), total_operations),
                'type': 'success',
                'sticky': False,
            }
        }

    def action_start_selected(self):
        """Start selected work orders and auto-issue available materials"""
        selected_lines = self.work_order_line_ids.filtered(lambda l: l.selected)

        if not selected_lines:
            raise UserError(_('Please select at least one work order to start!'))

        total_issued = 0
        total_moves = 0
        messages = []

        for line in selected_lines:
            line.action_start_production()
            issued, total = self._auto_issue_materials(line.production_id)
            total_issued += issued
            total_moves += total

            if issued > 0:
                messages.append(_('Production %s: %s/%s materials issued') %
                                (line.production_id.name, issued, total))

        self.state = 'in_progress'

        # Trigger project state update
        if self.project_id and self.project_id.auto_update_state:
            self.project_id.update_project_state()

        if total_issued > 0:
            message = _('%s work orders started!\n%s/%s materials auto-issued.\n\n') % (
                len(selected_lines), total_issued, total_moves
            )
            message += '\n'.join(messages)
            notification_type = 'success'
        else:
            message = _('%s work orders started!\nNo materials available to issue automatically.') % len(selected_lines)
            notification_type = 'warning'

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Work Orders Started'),
                'message': message,
                'type': notification_type,
                'sticky': True,
            }
        }

    def action_done(self):
        """Mark execution as done and update project"""
        self.write({'state': 'done'})

        # Trigger project state update
        for execution in self:
            if execution.project_id and execution.project_id.auto_update_state:
                execution.project_id.update_project_state()
    def _load_operations_for_line(self, execution_line):
        """Load operations for a production order into operation lines"""
        if not execution_line.production_id.workorder_ids:
            _logger.warning('No workorders found for production %s', execution_line.production_id.name)
            return 0

        operation_vals = []
        sequence = 10
        for workorder in execution_line.production_id.workorder_ids.sorted(lambda w: w.id):
            op_name = workorder.name
            if not op_name and workorder.operation_id:
                op_name = workorder.operation_id.name
            if not op_name:
                op_name = 'Operation %s' % workorder.id

            operation_vals.append({
                'execution_line_id': execution_line.id,
                'workorder_id': workorder.id,
                'name': op_name,
                'workcenter_id': workorder.workcenter_id.id if workorder.workcenter_id else False,
                'duration_expected': workorder.duration_expected or 0.0,
                'sequence': sequence,
            })
            sequence += 10

        if operation_vals:
            self.env['work.order.operation.line'].create(operation_vals)
            _logger.info('Created %d operation lines for %s', len(operation_vals), execution_line.component_id.name)
            return len(operation_vals)

        return 0

    def action_start_selected(self):
        """Start selected work orders and auto-issue available materials"""
        selected_lines = self.work_order_line_ids.filtered(lambda l: l.selected)

        if not selected_lines:
            raise UserError(_('Please select at least one work order to start!'))

        total_issued = 0
        total_moves = 0
        messages = []

        for line in selected_lines:
            line.action_start_production()
            issued, total = self._auto_issue_materials(line.production_id)
            total_issued += issued
            total_moves += total

            if issued > 0:
                messages.append(_('Production %s: %s/%s materials issued') %
                                (line.production_id.name, issued, total))

        self.state = 'in_progress'

        if total_issued > 0:
            message = _('%s work orders started!\n%s/%s materials auto-issued.\n\n') % (
                len(selected_lines), total_issued, total_moves
            )
            message += '\n'.join(messages)
            notification_type = 'success'
        else:
            message = _('%s work orders started!\nNo materials available to issue automatically.') % len(selected_lines)
            notification_type = 'warning'

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Work Orders Started'),
                'message': message,
                'type': notification_type,
                'sticky': True,
            }
        }

    def action_issue_materials_for_selected(self):
        """Manually issue available materials for selected work orders"""
        self.ensure_one()

        selected_lines = self.work_order_line_ids.filtered(lambda l: l.selected)

        if not selected_lines:
            raise UserError(_('Please select at least one work order first!'))

        total_issued = 0
        total_moves = 0
        messages = []
        stock_issues = []

        for line in selected_lines:
            production = line.production_id

            if production.state not in ('confirmed', 'progress', 'to_close'):
                if production.state == 'draft':
                    try:
                        production.action_confirm()
                        messages.append(_('ℹ️ Production %s: Confirmed automatically') % production.name)
                    except Exception as e:
                        messages.append(_('⚠️ Production %s: Cannot confirm - %s') % (production.name, str(e)))
                        continue
                else:
                    messages.append(_('⚠️ Production %s: Not in correct state (%s)') %
                                    (production.name, production.state))
                    continue

            raw_moves = production.move_raw_ids.filtered(lambda m: m.state not in ('done', 'cancel'))

            if not raw_moves:
                messages.append(_('ℹ️ Production %s: No materials to issue') % production.name)
                continue

            missing_materials = []
            for move in raw_moves:
                if move.product_id.qty_available < move.product_uom_qty:
                    missing_materials.append({
                        'product': move.product_id.name,
                        'required': move.product_uom_qty,
                        'available': move.product_id.qty_available,
                        'missing': move.product_uom_qty - move.product_id.qty_available
                    })

            try:
                production.action_assign()
            except Exception as e:
                _logger.warning('Could not assign production %s: %s', production.name, str(e))

            issued, total = self._auto_issue_materials(production)
            total_issued += issued
            total_moves += total

            if issued > 0:
                messages.append(_('✅ Production %s: %s/%s materials issued') %
                                (production.name, issued, total))
            elif total > 0:
                messages.append(_('⚠️ Production %s: 0/%s materials issued') %
                                (production.name, total))

                if missing_materials:
                    stock_msg = _('   Missing materials:')
                    for mat in missing_materials:
                        stock_msg += _('\n   • %s: Need %.2f, Have %.2f (Missing: %.2f)') % (
                            mat['product'], mat['required'], mat['available'], mat['missing']
                        )
                    stock_issues.append(stock_msg)
            else:
                messages.append(_('ℹ️ Production %s: No materials to issue') % production.name)

        if total_issued > 0:
            title = _('Materials Issued Successfully')
            message = _('✅ Issued %s out of %s materials.\n\n') % (total_issued, total_moves)
            message += '\n'.join(messages)
            if stock_issues:
                message += _('\n\n📋 Stock Issues:\n') + '\n'.join(stock_issues)
            notification_type = 'success'
        elif total_moves > 0:
            title = _('No Materials Available to Issue')
            message = _('⚠️ Could not issue materials. Check stock levels.\n\n')
            message += '\n'.join(messages)
            if stock_issues:
                message += _('\n\n📋 Stock Issues:\n') + '\n'.join(stock_issues)
            notification_type = 'warning'
        else:
            title = _('No Materials to Issue')
            message = _('ℹ️ Selected work orders have no pending materials.\n\n')
            message += '\n'.join(messages)
            notification_type = 'info'

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': title,
                'message': message,
                'type': notification_type,
                'sticky': True,
            }
        }

    def action_open_process_wizard(self):
        """Open the process all wizard"""
        self.ensure_one()

        selected_lines = self.work_order_line_ids.filtered(lambda l: l.selected)
        if not selected_lines:
            raise UserError(_('Please select at least one work order first!'))

        return {
            'type': 'ir.actions.act_window',
            'name': _('Process Work Orders'),
            'res_model': 'work.order.process.wizard',
            'view_mode': 'form',
            'target': 'new',
            'context': {
                'default_execution_id': self.id,
            }
        }

    def _auto_issue_materials(self, production):
        """Automatically issue available materials for a production order"""
        issued_count = 0
        total_count = 0

        if production.state == 'draft':
            try:
                production.action_confirm()
            except Exception as e:
                _logger.warning('Could not confirm production %s: %s', production.name, str(e))
                return 0, 0

        if production.state in ('confirmed', 'progress', 'to_close'):
            try:
                production.action_assign()
            except Exception as e:
                _logger.warning('Could not check availability for production %s: %s', production.name, str(e))

        raw_moves = production.move_raw_ids.filtered(
            lambda m: m.state not in ('done', 'cancel')
        )

        total_count = len(raw_moves)

        if total_count == 0:
            return 0, 0

        for move in raw_moves:
            try:
                if move.state in ('done', 'cancel'):
                    continue

                product = move.product_id
                required_qty = move.product_uom_qty
                available_qty = product.qty_available

                if move.state != 'assigned':
                    try:
                        move._action_assign()
                    except Exception as e:
                        _logger.warning('Could not reserve %s: %s', product.name, str(e))

                if move.state == 'assigned':
                    try:
                        if move.move_line_ids:
                            for move_line in move.move_line_ids:
                                if hasattr(move_line, 'quantity'):
                                    if move_line.quantity > 0:
                                        continue
                                    reserved = move_line.quantity if hasattr(move_line,
                                                                             'quantity') else move_line.reserved_qty
                                    move_line.quantity = reserved
                                else:
                                    if move_line.qty_done > 0:
                                        continue
                                    move_line.qty_done = move_line.reserved_qty
                        else:
                            move._action_assign()
                            for move_line in move.move_line_ids:
                                if hasattr(move_line, 'quantity'):
                                    reserved = move_line.quantity if hasattr(move_line,
                                                                             'quantity') else move_line.reserved_qty
                                    move_line.quantity = reserved
                                else:
                                    move_line.qty_done = move_line.reserved_qty

                        if move.move_line_ids:
                            if hasattr(move.move_line_ids[0], 'quantity'):
                                total_qty = sum(move.move_line_ids.mapped('quantity'))
                            else:
                                total_qty = sum(move.move_line_ids.mapped('qty_done'))

                            if total_qty > 0:
                                move._action_done()
                                issued_count += 1
                    except Exception as e:
                        _logger.warning('Could not validate move for %s: %s', product.name, str(e))

                elif available_qty >= required_qty:
                    try:
                        move._action_assign()

                        if move.state == 'assigned' and move.move_line_ids:
                            for move_line in move.move_line_ids:
                                if hasattr(move_line, 'quantity'):
                                    reserved = move_line.quantity if hasattr(move_line,
                                                                             'quantity') else move_line.reserved_qty
                                    move_line.quantity = reserved
                                else:
                                    move_line.qty_done = move_line.reserved_qty

                            move._action_done()
                            issued_count += 1
                    except Exception as e:
                        _logger.warning('Could not force-issue %s: %s', product.name, str(e))

            except Exception as e:
                _logger.error('Error processing material %s for production %s: %s',
                              move.product_id.name if move.product_id else 'Unknown',
                              production.name, str(e))
                continue

        return issued_count, total_count

    def action_done(self):
        self.write({'state': 'done'})

    def action_cancel(self):
        self.write({'state': 'cancelled'})

    def action_export_operations_excel(self):
        """Export operations tracking to Excel WITH SPECIFICATIONS AND COMPANY HEADER"""
        self.ensure_one()

        if not self.work_order_line_ids:
            raise UserError(_('No work orders loaded yet!'))

        all_operations = self.env['work.order.operation.line'].search([
            ('execution_id', '=', self.id)
        ])

        if not all_operations:
            raise UserError(_('No operations found!'))

        # Get planning for specifications
        planning = self.env['material.production.planning'].search([
            ('project_id', '=', self.project_id.id),
            ('product_id', '=', self.product_id.id),
        ], limit=1, order='create_date desc')

        operations_data = []
        for op_line in all_operations:
            component = op_line.component_id
            unit_weight = component.weight if component else 0
            quantity = op_line.execution_line_id.quantity if op_line.execution_line_id else 0
            qty_to_produce = op_line.qty_production or 0
            qty_produced = op_line.qty_produced or 0

            # Get specifications
            specs_text = ''
            if planning and component:
                planning_comp = planning.component_line_ids.filtered(
                    lambda c: c.component_id == component
                )
                if planning_comp and planning_comp.specification_ids:
                    specs = []
                    for spec in planning_comp.specification_ids.sorted(lambda s: s.sequence):
                        specs.append(f"{spec.specification_name}: {spec.value}")
                    specs_text = '; '.join(specs)

            operations_data.append({
                'production_order': op_line.production_id.name if op_line.production_id else '',
                'component': component.display_name if component else '',
                'specifications': specs_text,
                'unit_weight': unit_weight,
                'quantity': quantity,
                'total_weight': quantity * unit_weight,
                'operation': op_line.name or '',
                'workcenter': op_line.workcenter_id.name if op_line.workcenter_id else '',
                'state': dict(self.env['mrp.workorder']._fields['state'].selection).get(op_line.state,
                                                                                        op_line.state or 'pending'),
                'qty_to_produce': qty_to_produce,
                'weight_to_produce': qty_to_produce * unit_weight,
                'qty_produced': qty_produced,
                'weight_produced': qty_produced * unit_weight,
                'progress': op_line.progress_percentage or 0,
                'expected_duration': op_line.duration_expected or 0,
                'real_duration': op_line.duration_real or 0,
                'start_date': op_line.date_start.strftime('%Y-%m-%d %H:%M') if op_line.date_start else '',
                'finish_date': op_line.date_finished.strftime('%Y-%m-%d %H:%M') if op_line.date_finished else '',
            })

        if not operations_data:
            raise UserError(_('No operation data to export!'))

        try:
            import xlsxwriter

            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet('Operations Tracking')

            # Add company header with logo - FIXED VERSION
            row = ExcelExportHelper.add_company_header(
                workbook, worksheet, self.env,
                title=f"Operations Tracking - {self.name}",
                start_row=0
            )

            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4CAF50',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            cell_format = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'})
            number_format = workbook.add_format({'border': 1, 'align': 'right', 'num_format': '0.00'})
            percent_format = workbook.add_format({'border': 1, 'align': 'right', 'num_format': '0.00"%"'})
            text_wrap_format = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'top', 'text_wrap': True})

            status_formats = {
                'Done': workbook.add_format({'border': 1, 'bg_color': '#C8E6C9', 'align': 'center'}),
                'In Progress': workbook.add_format({'border': 1, 'bg_color': '#FFF9C4', 'align': 'center'}),
                'Progress': workbook.add_format({'border': 1, 'bg_color': '#FFF9C4', 'align': 'center'}),
                'Ready': workbook.add_format({'border': 1, 'bg_color': '#B3E5FC', 'align': 'center'}),
                'Pending': workbook.add_format({'border': 1, 'bg_color': '#F5F5F5', 'align': 'center'}),
                'Waiting': workbook.add_format({'border': 1, 'bg_color': '#F5F5F5', 'align': 'center'}),
                'Cancelled': workbook.add_format({'border': 1, 'bg_color': '#FFCDD2', 'align': 'center'}),
                'Cancel': workbook.add_format({'border': 1, 'bg_color': '#FFCDD2', 'align': 'center'}),
            }

            headers = [
                'Production Order', 'Component', 'Specifications', 'Unit Weight (kg)',
                'Quantity', 'Total Weight (kg)',
                'Operation', 'Work Center', 'State',
                'Qty to Produce', 'Weight to Produce (kg)',
                'Qty Produced', 'Weight Produced (kg)',
                'Progress %',
                'Expected Duration (min)', 'Real Duration (min)',
                'Start Date', 'Finish Date'
            ]

            for col, header in enumerate(headers):
                worksheet.write(row, col, header, header_format)

            row += 1
            for data in operations_data:
                col = 0
                worksheet.write(row, col, data['production_order'], cell_format)
                col += 1
                worksheet.write(row, col, data['component'], cell_format)
                col += 1
                worksheet.write(row, col, data['specifications'], text_wrap_format)
                col += 1
                worksheet.write(row, col, data['unit_weight'], number_format)
                col += 1
                worksheet.write(row, col, data['quantity'], number_format)
                col += 1
                worksheet.write(row, col, data['total_weight'], number_format)
                col += 1
                worksheet.write(row, col, data['operation'], cell_format)
                col += 1
                worksheet.write(row, col, data['workcenter'], cell_format)
                col += 1

                status = data['state']
                status_fmt = status_formats.get(status, cell_format)
                worksheet.write(row, col, status, status_fmt)
                col += 1

                worksheet.write(row, col, data['qty_to_produce'], number_format)
                col += 1
                worksheet.write(row, col, data['weight_to_produce'], number_format)
                col += 1
                worksheet.write(row, col, data['qty_produced'], number_format)
                col += 1
                worksheet.write(row, col, data['weight_produced'], number_format)
                col += 1
                worksheet.write(row, col, data['progress'], percent_format)
                col += 1
                worksheet.write(row, col, data['expected_duration'], number_format)
                col += 1
                worksheet.write(row, col, data['real_duration'], number_format)
                col += 1
                worksheet.write(row, col, data['start_date'], cell_format)
                col += 1
                worksheet.write(row, col, data['finish_date'], cell_format)
                col += 1

                row += 1

            worksheet.set_column('A:A', 20)
            worksheet.set_column('B:B', 25)
            worksheet.set_column('C:C', 40)  # Specifications
            worksheet.set_column('D:D', 15)
            worksheet.set_column('E:E', 12)
            worksheet.set_column('F:F', 16)
            worksheet.set_column('G:G', 20)
            worksheet.set_column('H:H', 15)
            worksheet.set_column('I:I', 15)
            worksheet.set_column('J:J', 14)
            worksheet.set_column('K:K', 18)
            worksheet.set_column('L:L', 14)
            worksheet.set_column('M:M', 18)
            worksheet.set_column('N:N', 12)
            worksheet.set_column('O:P', 18)
            worksheet.set_column('Q:R', 18)

            row += 2
            summary_header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4CAF50',
                'font_color': 'white',
                'border': 1,
                'align': 'left'
            })

            worksheet.write(row, 0, 'SUMMARY', summary_header_format)
            row += 1

            total_quantity = sum(d['quantity'] for d in operations_data)
            total_weight = sum(d['total_weight'] for d in operations_data)
            total_to_produce = sum(d['qty_to_produce'] for d in operations_data)
            total_weight_to_produce = sum(d['weight_to_produce'] for d in operations_data)
            total_produced = sum(d['qty_produced'] for d in operations_data)
            total_weight_produced = sum(d['weight_produced'] for d in operations_data)
            total_expected_duration = sum(d['expected_duration'] for d in operations_data)
            total_real_duration = sum(d['real_duration'] for d in operations_data)

            status_counts = {}
            for data in operations_data:
                status = data['state']
                status_counts[status] = status_counts.get(status, 0) + 1

            summary_data = [
                ('Total Operations:', len(operations_data), ''),
                ('', '', ''),
                ('Total Component Quantity:', total_quantity, ''),
                ('Total Component Weight:', total_weight, 'kg'),
                ('', '', ''),
                ('Total Qty to Produce:', total_to_produce, ''),
                ('Total Weight to Produce:', total_weight_to_produce, 'kg'),
                ('', '', ''),
                ('Total Qty Produced:', total_produced, ''),
                ('Total Weight Produced:', total_weight_produced, 'kg'),
                ('', '', ''),
                ('Total Expected Duration:', total_expected_duration, 'min'),
                ('Total Real Duration:', total_real_duration, 'min'),
                ('', '', ''),
            ]

            for label, value, unit in summary_data:
                worksheet.write(row, 0, label, cell_format)
                if value != '':
                    worksheet.write(row, 1, value, number_format)
                if unit:
                    worksheet.write(row, 2, unit, cell_format)
                row += 1

            worksheet.write(row, 0, 'Operations by Status:', cell_format)
            row += 1
            for status, count in sorted(status_counts.items()):
                worksheet.write(row, 0, f'  {status}:', cell_format)
                worksheet.write(row, 1, count, number_format)
                row += 1

            worksheet.freeze_panes(1, 0)

            workbook.close()
            output.seek(0)

            file_data = base64.b64encode(output.read())
            filename = 'Operations_Tracking_%s.xlsx' % self.name.replace('/', '_')

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

    def action_export_material_planning_excel(self):
        """Export material planning with detailed tracking AND COMPANY HEADER"""
        self.ensure_one()

        if not self.work_order_line_ids:
            raise UserError(_('No work orders loaded yet!'))

        materials_data = []

        for line in self.work_order_line_ids:
            production = line.production_id
            if not production:
                continue

            bom = production.bom_id
            if not bom:
                materials_data.append(self._get_material_data(
                    line.component_id,
                    line.component_id,
                    line.quantity,
                    production
                ))
                continue

            for bom_line in bom.bom_line_ids:
                material = bom_line.product_id
                required_qty = bom_line.product_qty * production.product_qty

                materials_data.append(self._get_material_data(
                    line.component_id,
                    material,
                    required_qty,
                    production
                ))

        if not materials_data:
            raise UserError(_('No material data found for export!'))

        try:
            import xlsxwriter

            output = io.BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet('Material Planning')

            # Add company header with logo - FIXED VERSION
            row = ExcelExportHelper.add_company_header(
                workbook, worksheet, self.env,
                title=f"Material Planning - {self.name}",
                start_row=0
            )

            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#2E86C1',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True
            })

            cell_format = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'})
            number_format = workbook.add_format({'border': 1, 'align': 'right', 'num_format': '0.00'})
            currency_format = workbook.add_format({'border': 1, 'align': 'right', 'num_format': '#,##0.00'})

            status_colors = {
                'sufficient': workbook.add_format({'border': 1, 'bg_color': '#D5F4E6', 'align': 'center'}),
                'partial': workbook.add_format({'border': 1, 'bg_color': '#FCF3CF', 'align': 'center'}),
                'shortage': workbook.add_format({'border': 1, 'bg_color': '#F9EBEA', 'align': 'center'}),
                'ordered': workbook.add_format({'border': 1, 'bg_color': '#D6EAF8', 'align': 'center'}),
            }

            headers = [
                'Production Order', 'Component', 'Material', 'UoM', 'Unit Weight (kg)',
                'Required Qty', 'Required Weight (kg)',
                'Available Qty', 'Available Weight (kg)',
                'Reserved Qty', 'Reserved Weight (kg)',
                'Free Qty', 'Free Weight (kg)',
                'Shortage Qty', 'Shortage Weight (kg)',
                'In RFQ Qty', 'In RFQ Weight (kg)',
                'In PO Qty', 'In PO Weight (kg)',
                'Received Qty', 'Received Weight (kg)',
                'Issued Qty', 'Issued Weight (kg)',
                'Status', 'Standard Cost', 'Total Cost'
            ]

            for col, header in enumerate(headers):
                worksheet.write(row, col, header, header_format)

            row += 1
            total_mat_weight = 0
            total_mat_cost = 0

            for data in materials_data:
                col = 0
                material_weight = data.get('unit_weight', 0)

                worksheet.write(row, col, data['production_order'], cell_format)
                col += 1
                worksheet.write(row, col, data['component'], cell_format)
                col += 1
                worksheet.write(row, col, data['material'], cell_format)
                col += 1
                worksheet.write(row, col, data['uom'], cell_format)
                col += 1
                worksheet.write(row, col, material_weight, number_format)
                col += 1

                worksheet.write(row, col, data['required_qty'], number_format)
                col += 1
                worksheet.write(row, col, data['required_qty'] * material_weight, number_format)
                col += 1

                worksheet.write(row, col, data['available_stock'], number_format)
                col += 1
                worksheet.write(row, col, data['available_stock'] * material_weight, number_format)
                col += 1

                worksheet.write(row, col, data['reserved_qty'], number_format)
                col += 1
                worksheet.write(row, col, data['reserved_qty'] * material_weight, number_format)
                col += 1

                worksheet.write(row, col, data['free_stock'], number_format)
                col += 1
                worksheet.write(row, col, data['free_stock'] * material_weight, number_format)
                col += 1

                worksheet.write(row, col, data['shortage_qty'], number_format)
                col += 1
                worksheet.write(row, col, data['shortage_qty'] * material_weight, number_format)
                col += 1

                worksheet.write(row, col, data['in_rfq_qty'], number_format)
                col += 1
                worksheet.write(row, col, data['in_rfq_qty'] * material_weight, number_format)
                col += 1

                worksheet.write(row, col, data['in_po_qty'], number_format)
                col += 1
                worksheet.write(row, col, data['in_po_qty'] * material_weight, number_format)
                col += 1

                worksheet.write(row, col, data['received_qty'], number_format)
                col += 1
                worksheet.write(row, col, data['received_qty'] * material_weight, number_format)
                col += 1

                worksheet.write(row, col, data['issued_qty'], number_format)
                col += 1
                worksheet.write(row, col, data['issued_qty'] * material_weight, number_format)
                col += 1

                status = data['status']
                status_fmt = status_colors.get(status, cell_format)
                worksheet.write(row, col, status.title(), status_fmt)
                col += 1

                worksheet.write(row, col, data['standard_cost'], currency_format)
                col += 1
                worksheet.write(row, col, data['total_cost'], currency_format)
                col += 1

                total_mat_weight += data['shortage_qty'] * material_weight
                total_mat_cost += data['total_cost']
                row += 1

            worksheet.write(row, 0, 'TOTAL', header_format)
            worksheet.write(row, 8, total_mat_weight, number_format)
            worksheet.write(row, 10, total_mat_cost, currency_format)

            worksheet.set_column('A:D', 25)
            worksheet.set_column('E:W', 15)
            worksheet.set_column('X:X', 12)
            worksheet.set_column('Y:Z', 14)

            worksheet.freeze_panes(1, 0)

            workbook.close()
            output.seek(0)

            file_data = base64.b64encode(output.read())
            filename = 'Material_Planning_%s.xlsx' % self.name.replace('/', '_')

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
            raise UserError(_('Please install xlsxwriter library'))
        except Exception as e:
            raise UserError(_('Error creating Excel file: %s') % str(e))

    def _get_material_data(self, component, material, required_qty, production):
        """Get detailed material data for Excel export"""

        available_stock = material.qty_available
        reserved_qty = material.outgoing_qty
        free_stock = available_stock - reserved_qty

        planning = self.env['material.production.planning'].search([
            ('project_id', '=', self.project_id.id),
            ('product_id', '=', self.product_id.id),
        ], limit=1, order='create_date desc')

        linked_po_ids = planning.rfq_ids.ids if planning else []

        issued_moves = self.env['stock.move'].search([
            '|',
            '&',
            ('raw_material_production_id', '=', production.id),
            ('product_id', '=', material.id),
            '&',
            ('production_id', '=', production.id),
            ('product_id', '=', material.id),
            ('state', '=', 'done')
        ])
        issued_qty = sum(issued_moves.mapped('product_uom_qty'))

        if issued_qty == 0:
            issued_moves_alt = self.env['stock.move'].search([
                ('product_id', '=', material.id),
                ('reference', 'ilike', production.name),
                ('location_dest_id.usage', '=', 'production'),
                ('state', '=', 'done')
            ])
            issued_qty = sum(issued_moves_alt.mapped('product_uom_qty'))

        remaining_required = max(0, required_qty - issued_qty)
        shortage_qty = max(0, remaining_required - free_stock)

        rfq_domain = [
            ('product_id', '=', material.id),
            ('order_id.state', 'in', ['draft', 'sent', 'to approve']),
            '|',
            ('order_id.origin', '=', production.name),
            '&',
            ('order_id.id', 'in', linked_po_ids),
            ('order_id.origin', 'ilike', production.name)
        ]

        rfq_lines = self.env['purchase.order.line'].search(rfq_domain)
        counted_rfq_lines = set()
        in_rfq_qty = 0
        for line in rfq_lines:
            if line.id not in counted_rfq_lines:
                in_rfq_qty += line.product_qty
                counted_rfq_lines.add(line.id)

        po_domain = [
            ('product_id', '=', material.id),
            ('order_id.state', 'in', ['purchase', 'done']),
            '|',
            ('order_id.origin', '=', production.name),
            '&',
            ('order_id.id', 'in', linked_po_ids),
            ('order_id.origin', 'ilike', production.name)
        ]

        po_lines = self.env['purchase.order.line'].search(po_domain)
        counted_po_lines = set()
        in_po_qty = 0
        received_qty = 0
        for line in po_lines:
            if line.id not in counted_po_lines:
                in_po_qty += line.product_qty
                received_qty += line.qty_received
                counted_po_lines.add(line.id)

        if free_stock >= remaining_required:
            status = 'sufficient'
        elif in_po_qty > 0 or received_qty > 0:
            status = 'ordered'
        elif free_stock > 0:
            status = 'partial'
        else:
            status = 'shortage'

        standard_cost = material.standard_price
        total_cost = required_qty * standard_cost
        unit_weight = material.weight or 0

        return {
            'production_order': production.name,
            'component': component.display_name,
            'material': material.display_name,
            'uom': material.uom_id.name,
            'unit_weight': unit_weight,
            'required_qty': required_qty,
            'available_stock': available_stock,
            'reserved_qty': reserved_qty,
            'free_stock': free_stock,
            'shortage_qty': shortage_qty,
            'in_rfq_qty': in_rfq_qty,
            'in_po_qty': in_po_qty,
            'received_qty': received_qty,
            'issued_qty': issued_qty,
            'status': status,
            'standard_cost': standard_cost,
            'total_cost': total_cost,
        }

    def action_open_operations_view(self):
        """Open operations view showing only incomplete operations"""
        self.ensure_one()

        total_ops = self.env['work.order.operation.line'].search_count([
            ('execution_id', '=', self.id),
        ])

        if total_ops == 0:
            raise UserError(_('No operations found!'))

        return {
            'type': 'ir.actions.act_window',
            'name': _('Work Order Operations'),
            'res_model': 'work.order.operation.line',
            'view_mode': 'tree,form',
            'domain': [
                ('execution_id', '=', self.id),
            ],
            'context': {
                'default_execution_id': self.id,
                'search_default_filter_not_completed': 1,
            },
            'target': 'current',
        }

    def action_open_production_report(self):
        """Open production progress report"""
        return {
            'type': 'ir.actions.act_window',
            'name': _('Production Progress Report'),
            'res_model': 'production.progress.report',
            'view_mode': 'pivot,graph,tree',
            'context': {
                'search_default_project_id': self.project_id.id,
                'search_default_product_id': self.product_id.id,
            },
            'target': 'current',
        }

    def action_view_material_issues(self):
        """View material issues for selected work orders"""
        self.ensure_one()

        selected_lines = self.work_order_line_ids.filtered(lambda l: l.selected)
        if not selected_lines:
            raise UserError(_('Please select at least one work order first!'))

        production_ids = selected_lines.mapped('production_id').ids

        moves = self.env['stock.move'].search([
            '|',
            ('raw_material_production_id', 'in', production_ids),
            '&',
            ('production_id', 'in', production_ids),
            ('location_dest_id.usage', '=', 'production')
        ])

        if not moves:
            move_lines = self.env['stock.move.line'].search([
                ('production_id', 'in', production_ids),
                ('location_dest_id.usage', '=', 'production')
            ])
            moves = move_lines.mapped('move_id')

        return {
            'type': 'ir.actions.act_window',
            'name': _('Material Issues - Selected Work Orders'),
            'res_model': 'stock.move',
            'view_mode': 'tree,form',
            'domain': [('id', 'in', moves.ids)],
            'context': {
                'default_location_id': self.env.ref('stock.stock_location_stock').id,
                'search_default_by_product': 1,
            },
            'target': 'current',
        }

    def action_view_operations_actual(self):
        """View actual operations (workorders) for selected work orders"""
        self.ensure_one()

        selected_lines = self.work_order_line_ids.filtered(lambda l: l.selected)
        if not selected_lines:
            raise UserError(_('Please select at least one work order first!'))

        production_ids = selected_lines.mapped('production_id').ids

        workorders = self.env['mrp.workorder'].search([
            ('production_id', 'in', production_ids),
        ])

        return {
            'type': 'ir.actions.act_window',
            'name': _('Operations Actual - Selected Work Orders'),
            'res_model': 'mrp.workorder',
            'view_mode': 'tree,form,gantt,pivot',
            'domain': [('id', 'in', workorders.ids)],
            'context': {
                'search_default_group_workcenter': 1,
            },
            'target': 'current',
        }


class WorkOrderExecutionLine(models.Model):
    _name = 'work.order.execution.line'
    _description = 'Work Order Execution Line'
    _order = 'sequence, id'

    sequence = fields.Integer(string='Sequence', default=10)
    execution_id = fields.Many2one(
        'work.order.execution',
        string='Execution',
        required=True,
        ondelete='cascade'
    )
    selected = fields.Boolean(
        string='Select',
        help='Select this line to execute'
    )
    component_id = fields.Many2one(
        'product.product',
        string='Component',
        required=True
    )
    quantity = fields.Float(
        string='Quantity',
        digits='Product Unit of Measure'
    )
    weight = fields.Float(
        string='Weight',
        digits='Stock Weight'
    )
    production_id = fields.Many2one(
        'mrp.production',
        string='Production Order',
        required=True
    )
    production_state = fields.Selection(
        related='production_id.state',
        string='Production State',
        store=True
    )

    operation_line_ids = fields.One2many(
        'work.order.operation.line',
        'execution_line_id',
        string='Operations'
    )

    current_operation = fields.Char(
        string='Current Operation',
        compute='_compute_current_operation'
    )
    progress_percentage = fields.Float(
        string='Progress %',
        compute='_compute_progress'
    )

    specifications_display = fields.Text(
        string='Specifications',
        compute='_compute_specifications_display',
        store=False,
        help='Display of all specifications from planning'
    )

    @api.depends('component_id', 'execution_id.project_id', 'execution_id.product_id')
    def _compute_specifications_display(self):
        """Get specifications from planning component"""
        for record in self:
            if record.execution_id and record.component_id:
                # Find the planning
                planning = self.env['material.production.planning'].search([
                    ('project_id', '=', record.execution_id.project_id.id),
                    ('product_id', '=', record.execution_id.product_id.id),
                ], limit=1, order='create_date desc')

                if planning:
                    # Find the component in planning
                    planning_comp = planning.component_line_ids.filtered(
                        lambda c: c.component_id == record.component_id
                    )

                    if planning_comp and planning_comp.specification_ids:
                        specs = []
                        for spec in planning_comp.specification_ids.sorted(lambda s: s.sequence):
                            specs.append(f"{spec.specification_name}: {spec.value}")
                        record.specifications_display = '\n'.join(specs)
                    else:
                        record.specifications_display = ''
                else:
                    record.specifications_display = ''
            else:
                record.specifications_display = ''

    @api.depends('production_id', 'production_id.workorder_ids', 'production_id.workorder_ids.state')
    def _compute_current_operation(self):
        for line in self:
            if line.production_id and line.production_id.workorder_ids:
                current_wo = line.production_id.workorder_ids.filtered(
                    lambda w: w.state in ('ready', 'progress')
                )
                if current_wo:
                    line.current_operation = current_wo[0].name
                else:
                    done_count = len(line.production_id.workorder_ids.filtered(lambda w: w.state == 'done'))
                    total_count = len(line.production_id.workorder_ids)
                    if done_count == total_count:
                        line.current_operation = _('All Operations Complete')
                    else:
                        line.current_operation = _('Not Started')
            else:
                line.current_operation = _('No Operations')

    @api.depends('production_id', 'production_id.workorder_ids.state')
    def _compute_progress(self):
        for line in self:
            if line.production_id and line.production_id.workorder_ids:
                total = len(line.production_id.workorder_ids)
                done = len(line.production_id.workorder_ids.filtered(
                    lambda w: w.state == 'done'
                ))
                line.progress_percentage = (done / total * 100) if total > 0 else 0
            else:
                line.progress_percentage = 0

    def action_start_production(self):
        """Start production order"""
        self.ensure_one()

        if self.production_id.state == 'draft':
            self.production_id.action_confirm()

        if self.production_id.state == 'confirmed':
            self.production_id.action_assign()

            if self.production_id.workorder_ids:
                first_wo = self.production_id.workorder_ids.filtered(
                    lambda w: w.state in ('pending', 'ready', 'waiting')
                )
                if first_wo:
                    first_wo[0].button_start()

    def action_next_operation(self):
        """Move to next operation"""
        self.ensure_one()

        if not self.production_id.workorder_ids:
            raise UserError(_('No work orders found for this production!'))

        current_wo = self.production_id.workorder_ids.filtered(
            lambda w: w.state == 'progress'
        )

        if current_wo:
            current_wo[0].button_finish()

            next_wo = self.production_id.workorder_ids.filtered(
                lambda w: w.state in ('ready', 'waiting')
            )
            if next_wo:
                next_wo[0].button_start()
        else:
            raise UserError(_('No work order in progress!'))

    def action_view_production(self):
        """View production order"""
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _('Production Order'),
            'res_model': 'mrp.production',
            'view_mode': 'form',
            'res_id': self.production_id.id,
            'target': 'current',
        }


class WorkOrderOperationLine(models.Model):
    _name = 'work.order.operation.line'
    _description = 'Work Order Operation Line'
    _order = 'sequence, id'

    sequence = fields.Integer(string='Sequence', default=10)
    execution_line_id = fields.Many2one(
        'work.order.execution.line',
        string='Execution Line',
        required=True,
        ondelete='cascade'
    )
    execution_id = fields.Many2one(
        related='execution_line_id.execution_id',
        string='Execution',
        store=True,
        index=True
    )
    component_id = fields.Many2one(
        related='execution_line_id.component_id',
        string='Component',
        store=True
    )
    production_id = fields.Many2one(
        related='execution_line_id.production_id',
        string='Production Order',
        store=True,
        index=True
    )
    workorder_id = fields.Many2one(
        'mrp.workorder',
        string='Work Order',
        index=True
    )
    name = fields.Char(string='Operation', required=True)
    operation_id = fields.Many2one(
        'mrp.routing.workcenter',
        string='Operation',
        related='workorder_id.operation_id',
        store=True
    )
    workcenter_id = fields.Many2one(
        'mrp.workcenter',
        string='Work Center'
    )
    state = fields.Selection(
        related='workorder_id.state',
        string='State',
        store=True
    )
    duration_expected = fields.Float(
        string='Expected Duration (minutes)'
    )
    duration_real = fields.Float(
        related='workorder_id.duration',
        string='Real Duration (minutes)'
    )
    qty_production = fields.Float(
        related='workorder_id.qty_production',
        string='Quantity to Produce',
        store=True
    )
    qty_produced = fields.Float(
        related='workorder_id.qty_produced',
        string='Quantity Produced',
        store=True
    )

    selected = fields.Boolean(
        string='Select',
        help='Select this operation to execute'
    )
    is_completed = fields.Boolean(
        string='Completed',
        compute='_compute_is_completed',
        store=True,
        index=True
    )
    progress_percentage = fields.Float(
        string='Progress %',
        compute='_compute_progress',
        store=True
    )
    date_start = fields.Datetime(
        related='workorder_id.date_start',
        string='Start Date',
        store=True
    )
    date_finished = fields.Datetime(
        related='workorder_id.date_finished',
        string='Finish Date',
        store=True
    )

    @api.depends('state')
    def _compute_is_completed(self):
        for record in self:
            record.is_completed = record.state in ('done', 'cancel')

    @api.depends('qty_production', 'qty_produced')
    def _compute_progress(self):
        for record in self:
            if record.qty_production:
                record.progress_percentage = (record.qty_produced / record.qty_production) * 100
            else:
                record.progress_percentage = 0.0

    def action_open_workorder(self):
        """Open the work order"""
        self.ensure_one()
        if not self.workorder_id:
            raise UserError(_('No work order linked to this operation!'))
        return {
            'type': 'ir.actions.act_window',
            'name': _('Work Order'),
            'res_model': 'mrp.workorder',
            'res_id': self.workorder_id.id,
            'view_mode': 'form',
            'target': 'current',
        }

    def action_start(self):
        """Start the work order"""
        for record in self:
            if record.workorder_id and record.state in ('pending', 'ready', 'waiting'):
                record.workorder_id.button_start()

    def action_finish(self):
        """Finish the work order"""
        for record in self:
            if record.workorder_id and record.state in ('progress', 'to_close'):
                record.workorder_id.button_finish()

    def action_start_selected(self):
        """Start all selected operations"""
        selected_ops = self.search([
            ('id', 'in', self.ids),
            ('selected', '=', True),
        ])

        if not selected_ops:
            raise UserError(_('Please select at least one operation to start!'))

        started_count = 0
        failed_count = 0
        messages = []

        for op_line in selected_ops:
            try:
                if op_line.state in ('pending', 'ready', 'waiting'):
                    if op_line.workorder_id:
                        op_line.workorder_id.button_start()
                        started_count += 1
                        messages.append(_('✅ %s: Started') % op_line.name)
                    else:
                        failed_count += 1
                        messages.append(_('⚠️ %s: No workorder linked') % op_line.name)
                elif op_line.state == 'progress':
                    messages.append(_('ℹ️ %s: Already in progress') % op_line.name)
                elif op_line.state == 'done':
                    messages.append(_('ℹ️ %s: Already completed') % op_line.name)
                else:
                    failed_count += 1
                    messages.append(_('⚠️ %s: Cannot start (state: %s)') % (op_line.name, op_line.state))
            except Exception as e:
                failed_count += 1
                messages.append(_('❌ %s: Error - %s') % (op_line.name, str(e)))
                _logger.error('Error starting operation %s: %s', op_line.name, str(e))

        if started_count > 0:
            title = _('Operations Started')
            message = _('✅ Started %s out of %s operations.\n\n') % (started_count, len(selected_ops))
            message += '\n'.join(messages[:10])
            if len(messages) > 10:
                message += _('\n... and %s more (check details)') % (len(messages) - 10)
            notification_type = 'success'
        elif failed_count > 0:
            title = _('Could Not Start Operations')
            message = _('⚠️ Failed to start operations.\n\n')
            message += '\n'.join(messages[:10])
            if len(messages) > 10:
                message += _('\n... and %s more') % (len(messages) - 10)
            notification_type = 'warning'
        else:
            title = _('No Operations to Start')
            message = _('ℹ️ Selected operations are already started or completed.\n\n')
            message += '\n'.join(messages[:10])
            notification_type = 'info'

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': title,
                'message': message,
                'type': notification_type,
                'sticky': True,
            }
        }

    def action_finish_selected(self):
        """Finish all selected operations"""
        selected_ops = self.search([
            ('id', 'in', self.ids),
            ('selected', '=', True),
        ])

        if not selected_ops:
            raise UserError(_('Please select at least one operation to finish!'))

        finished_count = 0
        failed_count = 0
        messages = []

        for op_line in selected_ops:
            try:
                if op_line.state in ('progress', 'to_close'):
                    if op_line.workorder_id:
                        op_line.workorder_id.button_finish()
                        finished_count += 1
                        messages.append(_('✅ %s: Finished') % op_line.name)
                    else:
                        failed_count += 1
                        messages.append(_('⚠️ %s: No workorder linked') % op_line.name)
                elif op_line.state == 'done':
                    messages.append(_('ℹ️ %s: Already completed') % op_line.name)
                elif op_line.state in ('pending', 'ready', 'waiting'):
                    messages.append(_('ℹ️ %s: Not started yet') % op_line.name)
                else:
                    failed_count += 1
                    messages.append(_('⚠️ %s: Cannot finish (state: %s)') % (op_line.name, op_line.state))
            except Exception as e:
                failed_count += 1
                messages.append(_('❌ %s: Error - %s') % (op_line.name, str(e)))
                _logger.error('Error finishing operation %s: %s', op_line.name, str(e))

        if finished_count > 0:
            title = _('Operations Finished')
            message = _('✅ Finished %s out of %s operations.\n\n') % (finished_count, len(selected_ops))
            message += '\n'.join(messages[:10])
            if len(messages) > 10:
                message += _('\n... and %s more') % (len(messages) - 10)
            notification_type = 'success'
        elif failed_count > 0:
            title = _('Could Not Finish Operations')
            message = _('⚠️ Failed to finish operations.\n\n')
            message += '\n'.join(messages[:10])
            if len(messages) > 10:
                message += _('\n... and %s more') % (len(messages) - 10)
            notification_type = 'warning'
        else:
            title = _('No Operations to Finish')
            message = _('ℹ️ Selected operations are not in progress.\n\n')
            message += '\n'.join(messages[:10])
            notification_type = 'info'

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': title,
                'message': message,
                'type': notification_type,
                'sticky': True,
            }
        }

    def action_process_selected_operations(self):
        """Process all selected operations"""
        selected_ops = self.search([
            ('id', 'in', self.ids),
            ('selected', '=', True),
        ])

        if not selected_ops:
            raise UserError(_('Please select at least one operation to process!'))

        execution_ids = selected_ops.mapped('execution_id')

        if len(execution_ids) > 1:
            raise UserError(_('Selected operations must belong to the same execution!'))

        execution = execution_ids[0]

        production_ids = selected_ops.mapped('production_id').ids

        wizard = self.env['work.order.process.wizard'].create({
            'execution_id': execution.id,
            'production_ids': [(6, 0, production_ids)],
        })

        wizard.operation_line_ids.unlink()
        wizard.material_line_ids.unlink()

        operation_lines = []
        for op_line in selected_ops:
            if op_line.workorder_id and op_line.state not in ('done', 'cancel'):
                operation_lines.append((0, 0, {
                    'production_id': op_line.production_id.id,
                    'workorder_id': op_line.workorder_id.id,
                    'operation_id': op_line.operation_id.id if op_line.operation_id else False,
                    'workcenter_id': op_line.workcenter_id.id if op_line.workcenter_id else False,
                    'qty_to_produce': op_line.qty_production,
                    'qty_produced': op_line.qty_produced,
                    'qty_remaining': op_line.qty_production - op_line.qty_produced,
                    'duration_expected': op_line.duration_expected,
                    'duration_hours': op_line.duration_expected / 60.0 if op_line.duration_expected else 0,
                    'state': op_line.state,
                }))

        material_lines = []
        for production in selected_ops.mapped('production_id'):
            raw_moves = production.move_raw_ids.filtered(
                lambda m: m.state not in ('done', 'cancel')
            )

            for move in raw_moves:
                material_lines.append((0, 0, {
                    'production_id': production.id,
                    'product_id': move.product_id.id,
                    'move_id': move.id,
                    'required_qty': move.product_uom_qty,
                    'available_qty': move.product_id.qty_available,
                    'qty_to_issue': min(move.product_uom_qty, move.product_id.qty_available),
                    'uom_id': move.product_uom.id,
                    'unit_weight': move.product_id.weight or 0,
                }))

        wizard.write({
            'operation_line_ids': operation_lines,
            'material_line_ids': material_lines,
        })

        return {
            'type': 'ir.actions.act_window',
            'name': _('Process Selected Operations'),
            'res_model': 'work.order.process.wizard',
            'res_id': wizard.id,
            'view_mode': 'form',
            'target': 'new',
        }


class WorkOrderOperationReport(models.Model):
    """Report model for operation-based tracking"""
    _name = 'work.order.operation.report'
    _description = 'Work Order Operation Report'
    _auto = False
    _order = 'production_name, component_name, sequence'

    id = fields.Integer(string='ID', readonly=True)
    execution_id = fields.Many2one('work.order.execution', string='Execution', readonly=True)
    production_id = fields.Many2one('mrp.production', string='Production Order', readonly=True)
    production_name = fields.Char(string='Production Order', readonly=True)
    component_id = fields.Many2one('product.product', string='Component', readonly=True)
    component_name = fields.Char(string='Component', readonly=True)
    workorder_id = fields.Many2one('mrp.workorder', string='Work Order', readonly=True)
    operation_name = fields.Char(string='Operation', readonly=True)
    workcenter_name = fields.Char(string='Work Center', readonly=True)
    sequence = fields.Integer(string='Sequence', readonly=True)
    state = fields.Selection([
        ('pending', 'Pending'),
        ('waiting', 'Waiting'),
        ('ready', 'Ready'),
        ('progress', 'In Progress'),
        ('done', 'Done'),
        ('cancel', 'Cancelled')
    ], string='State', readonly=True)
    qty_production = fields.Float(string='Quantity', readonly=True)
    qty_produced = fields.Float(string='Produced', readonly=True)
    progress_percentage = fields.Float(string='Progress %', readonly=True)
    duration_expected = fields.Float(string='Expected Duration', readonly=True)
    duration_real = fields.Float(string='Real Duration', readonly=True)
    date_start = fields.Datetime(string='Start Date', readonly=True)
    date_finished = fields.Datetime(string='Finish Date', readonly=True)

    def init(self):
        tools.drop_view_if_exists(self._cr, 'work_order_operation_report')
        self._cr.execute("""
            CREATE OR REPLACE VIEW work_order_operation_report AS (
                SELECT 
                    wol.id as id,
                    wol.execution_id as execution_id,
                    wel.production_id as production_id,
                    mp.name as production_name,
                    wel.component_id as component_id,
                    COALESCE(pp.default_code, pt.name::text, 'Component') as component_name,
                    wol.workorder_id as workorder_id,
                    wol.name as operation_name,
                    mw.name as workcenter_name,
                    wol.sequence as sequence,
                    COALESCE(mwo.state, 'pending') as state,
                    0.0 as qty_production,
                    0.0 as qty_produced,
                    0.0 as progress_percentage,
                    COALESCE(wol.duration_expected, 0) as duration_expected,
                    0.0 as duration_real,
                    mwo.date_start as date_start,
                    mwo.date_finished as date_finished
                FROM work_order_operation_line wol
                LEFT JOIN work_order_execution_line wel ON wol.execution_line_id = wel.id
                LEFT JOIN mrp_production mp ON wel.production_id = mp.id
                LEFT JOIN product_product pp ON wel.component_id = pp.id
                LEFT JOIN product_template pt ON pp.product_tmpl_id = pt.id
                LEFT JOIN mrp_workcenter mw ON wol.workcenter_id = mw.id
                LEFT JOIN mrp_workorder mwo ON wol.workorder_id = mwo.id
            )
        """)