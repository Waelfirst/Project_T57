# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import UserError
import logging

_logger = logging.getLogger(__name__)


class WorkOrderProcessWizard(models.TransientModel):
    _name = 'work.order.process.wizard'
    _description = 'Process Work Orders Wizard'

    execution_id = fields.Many2one('work.order.execution', string='Execution', required=True, ondelete='cascade')
    production_ids = fields.Many2many('mrp.production', string='Selected Productions')

    material_line_ids = fields.One2many('work.order.process.material.line', 'wizard_id', string='Materials to Issue')
    operation_line_ids = fields.One2many('work.order.process.operation.line', 'wizard_id',
                                         string='Operations to Complete')

    # Summary fields
    total_materials = fields.Integer(compute='_compute_totals', string='Total Materials')
    total_operations = fields.Integer(compute='_compute_totals', string='Total Operations')
    total_material_weight = fields.Float(compute='_compute_totals', string='Total Material Weight (kg)')
    total_operation_hours = fields.Float(compute='_compute_totals', string='Total Operation Hours')

    @api.depends('material_line_ids', 'operation_line_ids')
    def _compute_totals(self):
        for wizard in self:
            wizard.total_materials = len(wizard.material_line_ids)
            wizard.total_operations = len(wizard.operation_line_ids)
            wizard.total_material_weight = sum(wizard.material_line_ids.mapped('total_weight'))
            wizard.total_operation_hours = sum(wizard.operation_line_ids.mapped('duration_hours'))

    @api.model
    def default_get(self, fields_list):
        """Populate wizard with materials and operations from selected work orders"""
        res = super(WorkOrderProcessWizard, self).default_get(fields_list)

        execution_id = self.env.context.get('active_id') or self.env.context.get('default_execution_id')
        if not execution_id:
            raise UserError(_('No work order execution found!'))

        execution = self.env['work.order.execution'].browse(execution_id)

        # Get selected work order lines
        selected_lines = execution.work_order_line_ids.filtered(lambda l: l.selected)
        if not selected_lines:
            raise UserError(_('Please select at least one work order first!'))

        production_ids = selected_lines.mapped('production_id').ids
        res['execution_id'] = execution_id
        res['production_ids'] = [(6, 0, production_ids)]

        # Collect materials
        material_lines = []
        for production in selected_lines.mapped('production_id'):
            # Get pending material moves
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

        res['material_line_ids'] = material_lines

        # Collect operations
        operation_lines = []
        for production in selected_lines.mapped('production_id'):
            # Get pending workorders
            workorders = production.workorder_ids.filtered(
                lambda w: w.state not in ('done', 'cancel')
            )

            for workorder in workorders:
                operation_lines.append((0, 0, {
                    'production_id': production.id,
                    'workorder_id': workorder.id,
                    'operation_id': workorder.operation_id.id if workorder.operation_id else False,
                    'workcenter_id': workorder.workcenter_id.id if workorder.workcenter_id else False,
                    'qty_to_produce': workorder.qty_production,
                    'qty_produced': workorder.qty_produced,
                    'qty_remaining': workorder.qty_remaining,
                    'duration_expected': workorder.duration_expected,
                    'duration_hours': workorder.duration_expected / 60.0 if workorder.duration_expected else 0,
                    'state': workorder.state,
                }))

        res['operation_line_ids'] = operation_lines

        return res

    def action_process_all(self):
        """Process all materials and operations"""
        self.ensure_one()

        materials_issued = 0
        materials_failed = 0
        operations_completed = 0
        operations_failed = 0
        errors = []

        _logger.info('=' * 80)
        _logger.info('Starting Process All for Execution: %s', self.execution_id.name)
        _logger.info('Materials to process: %s', len(self.material_line_ids))
        _logger.info('Operations to process: %s', len(self.operation_line_ids))
        _logger.info('=' * 80)

        # Process materials first
        for material_line in self.material_line_ids:
            material_name = material_line.product_id.name
            production_name = material_line.production_id.name

            try:
                if material_line.qty_to_issue > 0:
                    _logger.info('Processing material: %s for production: %s', material_name, production_name)

                    # Use savepoint to isolate each operation
                    try:
                        with self.env.cr.savepoint():
                            success = material_line._issue_material()
                            if success:
                                materials_issued += 1
                                _logger.info('✅ Successfully issued material: %s', material_name)
                            else:
                                materials_failed += 1
                                error_msg = _('Material %s: Failed to issue') % material_name
                                errors.append(error_msg)
                                _logger.warning('⚠️ %s', error_msg)
                    except Exception as e:
                        materials_failed += 1
                        error_msg = _('Material %s: %s') % (material_name, str(e))
                        errors.append(error_msg)
                        _logger.error('❌ Error in savepoint for material %s: %s', material_name, str(e))
                else:
                    _logger.info('Skipping material %s (qty_to_issue = 0)', material_name)

            except Exception as e:
                materials_failed += 1
                error_msg = _('Material %s: %s') % (material_name, str(e))
                errors.append(error_msg)
                _logger.error('❌ Outer error for material %s: %s', material_name, str(e))

        _logger.info('-' * 80)
        _logger.info('Materials processed: %s issued, %s failed', materials_issued, materials_failed)
        _logger.info('-' * 80)

        # Commit material issues before processing operations
        try:
            self.env.cr.commit()
            _logger.info('✅ Committed material transactions')
        except Exception as e:
            _logger.error('❌ Failed to commit material transactions: %s', str(e))

        # Process operations
        for operation_line in self.operation_line_ids:
            # Get operation name safely
            operation_name = 'Unknown Operation'
            production_name = 'Unknown Production'

            try:
                if operation_line.operation_id:
                    operation_name = operation_line.operation_id.name
                elif operation_line.workorder_id:
                    operation_name = operation_line.workorder_id.name or 'Unnamed Workorder'

                if operation_line.production_id:
                    production_name = operation_line.production_id.name
            except:
                pass

            try:
                if operation_line.state not in ('done', 'cancel'):
                    _logger.info('Processing operation: %s for production: %s', operation_name, production_name)

                    # Use savepoint to isolate each operation
                    try:
                        with self.env.cr.savepoint():
                            success = operation_line._complete_operation()
                            if success:
                                operations_completed += 1
                                _logger.info('✅ Successfully completed operation: %s', operation_name)
                            else:
                                operations_failed += 1
                                error_msg = _('Operation %s: Failed to complete') % operation_name
                                errors.append(error_msg)
                                _logger.warning('⚠️ %s', error_msg)
                    except Exception as e:
                        operations_failed += 1
                        error_msg = _('Operation %s: %s') % (operation_name, str(e))
                        errors.append(error_msg)
                        _logger.error('❌ Error in savepoint for operation %s: %s', operation_name, str(e))
                else:
                    _logger.info('Skipping operation %s (state: %s)', operation_name, operation_line.state)

            except Exception as e:
                operations_failed += 1
                error_msg = _('Operation %s: %s') % (operation_name, str(e))
                errors.append(error_msg)
                _logger.error('❌ Outer error for operation %s: %s', operation_name, str(e))

        _logger.info('-' * 80)
        _logger.info('Operations processed: %s completed, %s failed', operations_completed, operations_failed)
        _logger.info('=' * 80)

        # Prepare result message
        total_materials = len(self.material_line_ids)
        total_operations = len(self.operation_line_ids)

        if materials_issued > 0 or operations_completed > 0:
            message = _('✅ Processing Complete!\n\n')
            message += _('📦 Materials: %s issued, %s failed (out of %s)\n') % (
                materials_issued, materials_failed, total_materials
            )
            message += _('⚙️ Operations: %s completed, %s failed (out of %s)\n') % (
                operations_completed, operations_failed, total_operations
            )

            if errors:
                message += _('\n⚠️ Errors (%s total):\n') % len(errors)
                # Show first 10 errors
                for error in errors[:10]:
                    message += f'  • {error}\n'
                if len(errors) > 10:
                    message += _('\n... and %s more errors (check server logs)') % (len(errors) - 10)

            notification_type = 'success' if not errors else 'warning'
        else:
            message = _('⚠️ Nothing was processed!\n\n')
            message += _('📦 Materials: 0/%s processed\n') % total_materials
            message += _('⚙️ Operations: 0/%s processed\n') % total_operations

            if errors:
                message += _('\n❌ Errors (%s total):\n') % len(errors)
                for error in errors[:10]:
                    message += f'  • {error}\n'
                if len(errors) > 10:
                    message += _('\n... and %s more errors (check server logs)') % (len(errors) - 10)
            else:
                message += _('\nℹ️ No materials or operations to process.')

            notification_type = 'warning'

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Process Complete'),
                'message': message,
                'type': notification_type,
                'sticky': True,
            }
        }


class WorkOrderProcessMaterialLine(models.TransientModel):
    _name = 'work.order.process.material.line'
    _description = 'Process Material Line'

    wizard_id = fields.Many2one('work.order.process.wizard', string='Wizard', required=True, ondelete='cascade')
    production_id = fields.Many2one('mrp.production', string='Production Order', required=True)
    product_id = fields.Many2one('product.product', string='Material', required=True)
    move_id = fields.Many2one('stock.move', string='Stock Move')

    required_qty = fields.Float(string='Required Qty', digits='Product Unit of Measure')
    available_qty = fields.Float(string='Available Qty', digits='Product Unit of Measure')
    qty_to_issue = fields.Float(string='Qty to Issue', digits='Product Unit of Measure')

    uom_id = fields.Many2one('uom.uom', string='UoM', required=True)
    unit_weight = fields.Float(string='Unit Weight (kg)', digits='Stock Weight')
    total_weight = fields.Float(compute='_compute_total_weight', string='Total Weight (kg)', digits='Stock Weight',
                                store=True)

    can_issue = fields.Boolean(compute='_compute_can_issue', string='Can Issue')

    @api.depends('qty_to_issue', 'unit_weight')
    def _compute_total_weight(self):
        for line in self:
            line.total_weight = line.qty_to_issue * line.unit_weight

    @api.depends('qty_to_issue', 'available_qty')
    def _compute_can_issue(self):
        for line in self:
            line.can_issue = line.qty_to_issue > 0 and line.qty_to_issue <= line.available_qty

    def _issue_material(self):
        """Issue the material - FIXED FOR ODOO 17"""
        self.ensure_one()

        if not self.move_id:
            _logger.warning('No stock move found for material %s', self.product_id.name)
            return False

        move = self.move_id

        try:
            _logger.info('  Issuing %s units of %s', self.qty_to_issue, self.product_id.name)
            _logger.info('  Move state: %s', move.state)

            # Ensure move is in correct state
            if move.state == 'draft':
                _logger.info('  Confirming draft move...')
                move._action_confirm()

            # Try to reserve
            if move.state != 'assigned':
                _logger.info('  Attempting to reserve stock...')
                move._action_assign()
                _logger.info('  After reserve, move state: %s', move.state)

            # Set quantity done
            if move.state == 'assigned':
                _logger.info('  Move is assigned, setting quantities...')
                # Update or create move lines
                if move.move_line_ids:
                    _logger.info('  Found %s move lines', len(move.move_line_ids))
                    qty_remaining = self.qty_to_issue
                    for move_line in move.move_line_ids:
                        if qty_remaining <= 0:
                            break

                        # Get reserved quantity - FIXED FOR ODOO 17
                        if hasattr(move_line, 'quantity'):
                            reserved = move_line.quantity  # Odoo 17
                        elif hasattr(move_line, 'reserved_uom_qty'):
                            reserved = move_line.reserved_uom_qty  # Odoo 17 alternative
                        elif hasattr(move_line, 'product_uom_qty'):
                            reserved = move_line.product_uom_qty  # Fallback
                        else:
                            reserved = move_line.reserved_qty  # Older versions

                        qty_to_set = min(qty_remaining, reserved)
                        _logger.info('  Setting move_line quantity to %s (reserved: %s)',
                                     qty_to_set, reserved)

                        # ODOO 17 FIX: Use 'quantity' field instead of 'qty_done'
                        if hasattr(move_line, 'quantity'):
                            move_line.quantity = qty_to_set  # Odoo 17
                        else:
                            move_line.qty_done = qty_to_set  # Older versions

                        qty_remaining -= qty_to_set
                else:
                    _logger.info('  No move lines found, trying to assign again...')
                    move._action_assign()
                    for move_line in move.move_line_ids:
                        # Get reserved quantity - FIXED FOR ODOO 17
                        if hasattr(move_line, 'quantity'):
                            reserved = move_line.quantity
                        elif hasattr(move_line, 'reserved_uom_qty'):
                            reserved = move_line.reserved_uom_qty
                        elif hasattr(move_line, 'product_uom_qty'):
                            reserved = move_line.product_uom_qty
                        else:
                            reserved = move_line.reserved_qty

                        _logger.info('  Setting move_line quantity to reserved qty: %s', reserved)

                        # ODOO 17 FIX: Use 'quantity' field instead of 'qty_done'
                        if hasattr(move_line, 'quantity'):
                            move_line.quantity = reserved  # Odoo 17
                        else:
                            move_line.qty_done = reserved  # Older versions

                # Validate the move
                # ODOO 17 FIX: Read 'quantity' field instead of 'qty_done'
                if move.move_line_ids and hasattr(move.move_line_ids[0], 'quantity'):
                    total_qty_done = sum(move.move_line_ids.mapped('quantity'))  # Odoo 17
                else:
                    total_qty_done = sum(move.move_line_ids.mapped('qty_done'))  # Older versions

                _logger.info('  Total quantity set: %s', total_qty_done)

                if total_qty_done > 0:
                    _logger.info('  Validating move...')
                    move._action_done()
                    _logger.info('  Move validated successfully')
                    return True
                else:
                    _logger.warning('  No quantity set, cannot validate')
                    return False
            else:
                _logger.warning('  Move not in assigned state: %s', move.state)
                return False

        except Exception as e:
            _logger.error('Error issuing material %s: %s', self.product_id.name, str(e))
            _logger.exception('Full traceback:')
            raise


class WorkOrderProcessOperationLine(models.TransientModel):
    _name = 'work.order.process.operation.line'
    _description = 'Process Operation Line'

    wizard_id = fields.Many2one('work.order.process.wizard', string='Wizard', required=True, ondelete='cascade')
    production_id = fields.Many2one('mrp.production', string='Production Order', required=True)
    workorder_id = fields.Many2one('mrp.workorder', string='Work Order')
    operation_id = fields.Many2one('mrp.routing.workcenter', string='Operation')
    workcenter_id = fields.Many2one('mrp.workcenter', string='Work Center')

    qty_to_produce = fields.Float(string='Qty to Produce', digits='Product Unit of Measure')
    qty_produced = fields.Float(string='Qty Produced', digits='Product Unit of Measure')
    qty_remaining = fields.Float(string='Qty Remaining', digits='Product Unit of Measure')

    duration_expected = fields.Float(string='Expected Duration (min)')
    duration_hours = fields.Float(string='Duration (hours)', digits=(16, 2))
    duration_minutes = fields.Float(compute='_compute_duration_minutes', string='Duration (min)', store=True)

    state = fields.Selection([
        ('pending', 'Pending'),
        ('ready', 'Ready'),
        ('waiting', 'Waiting'),
        ('progress', 'In Progress'),
        ('done', 'Done'),
        ('cancel', 'Cancelled')
    ], string='State')

    can_complete = fields.Boolean(compute='_compute_can_complete', string='Can Complete')

    @api.depends('duration_hours')
    def _compute_duration_minutes(self):
        for line in self:
            line.duration_minutes = line.duration_hours * 60.0

    @api.depends('state')
    def _compute_can_complete(self):
        for line in self:
            line.can_complete = line.state not in ('done', 'cancel')

    def _complete_operation(self):
        """Complete the operation"""
        self.ensure_one()

        if not self.workorder_id:
            _logger.warning('No workorder found')
            return False

        workorder = self.workorder_id

        try:
            _logger.info('  Completing workorder: %s', workorder.name)
            _logger.info('  Current state: %s', workorder.state)

            # Start if not started
            if workorder.state in ('pending', 'ready', 'waiting'):
                _logger.info('  Starting workorder...')
                workorder.button_start()
                _logger.info('  Workorder started, new state: %s', workorder.state)

            # Set quantities
            if workorder.qty_producing == 0:
                _logger.info('  Setting qty_producing to %s', self.qty_to_produce)
                workorder.qty_producing = self.qty_to_produce

            # Record time if specified
            if self.duration_hours > 0 and self.workcenter_id:
                _logger.info('  Recording time: %s hours (%s minutes)',
                             self.duration_hours, self.duration_minutes)

                # Get or create a default "productive" loss type
                productive_loss = self.env['mrp.workcenter.productivity.loss'].search([
                    ('loss_type', '=', 'productive')
                ], limit=1)

                if not productive_loss:
                    _logger.warning('  No productive loss type found, creating one...')
                    productive_loss = self.env['mrp.workcenter.productivity.loss'].create({
                        'name': 'Productive Time',
                        'loss_type': 'productive',
                    })
                    _logger.info('  Created productive loss type: %s', productive_loss.id)

                # Create time tracking entry with required loss_id
                productivity = self.env['mrp.workcenter.productivity'].create({
                    'workorder_id': workorder.id,
                    'workcenter_id': self.workcenter_id.id,
                    'loss_id': productive_loss.id,  # Required field in Odoo 17
                    'description': _('Processed via wizard'),
                    'date_start': fields.Datetime.now(),
                    'date_end': fields.Datetime.now(),
                    'duration': self.duration_minutes,
                    'user_id': self.env.user.id,
                })
                _logger.info('  Time recorded successfully (productivity id: %s)', productivity.id)

            # Complete the workorder
            if workorder.state in ('progress', 'to_close'):
                _logger.info('  Finishing workorder...')
                workorder.button_finish()
                _logger.info('  Workorder finished successfully')
                return True
            else:
                _logger.warning('  Workorder not in correct state to finish: %s', workorder.state)
                return False

        except Exception as e:
            _logger.error('Error completing operation: %s', str(e))
            _logger.exception('Full traceback:')
            raise