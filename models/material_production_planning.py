# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import UserError


class MaterialProductionPlanning(models.Model):
    _name = 'material.production.planning'
    _description = 'Material & Production Planning'
    _inherit = ['mail.thread', 'mail.activity.mixin']
    _order = 'create_date desc'

    name = fields.Char(
        string='Planning Reference',
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
    pricing_id = fields.Many2one(
        'project.product.pricing',
        string='Pricing Reference',
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
        ('components_loaded', 'Components Loaded'),
        ('material_planned', 'Material Planned'),
        ('work_orders_created', 'Work Orders Created'),
        ('done', 'Done'),
        ('cancelled', 'Cancelled'),
    ], string='Status', default='draft', tracking=True)

    component_line_ids = fields.One2many(
        'material.planning.component',
        'planning_id',
        string='Component Lines'
    )

    material_requirement_ids = fields.One2many(
        'material.requirement.line',
        'planning_id',
        string='Material Requirements'
    )

    production_order_ids = fields.Many2many(
        'mrp.production',
        'material_planning_production_rel',
        'planning_id',
        'production_id',
        string='Production Orders',
        readonly=True
    )

    production_count = fields.Integer(
        string='Production Orders',
        compute='_compute_production_count'
    )

    total_produced_qty = fields.Float(
        string='Total Produced Quantity',
        compute='_compute_produced_quantities',
        store=True
    )

    remaining_qty = fields.Float(
        string='Remaining to Produce',
        compute='_compute_produced_quantities',
        store=True
    )

    work_order_ids = fields.Many2many(
        'mrp.workorder',
        string='Work Orders',
        readonly=True
    )

    rfq_ids = fields.Many2many(
        'purchase.order',
        string='RFQs/Purchase Orders',
        readonly=True
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
            vals['name'] = self.env['ir.sequence'].next_by_code('material.production.planning') or _('New')
        return super(MaterialProductionPlanning, self).create(vals)

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
        """Trigger project state update when planning is created"""
        if vals.get('name', _('New')) == _('New'):
            vals['name'] = self.env['ir.sequence'].next_by_code('material.production.planning') or _('New')

        planning = super(MaterialProductionPlanning, self).create(vals)

        # Update project state
        if planning.project_id and planning.project_id.auto_update_state:
            planning.project_id.update_project_state()

        return planning

    def write(self, vals):
        """Trigger project state update when planning state changes"""
        result = super(MaterialProductionPlanning, self).write(vals)

        # Update project state if state changed
        if 'state' in vals:
            for planning in self:
                if planning.project_id and planning.project_id.auto_update_state:
                    planning.project_id.update_project_state()

        return result

    def action_load_components(self):
        """Update state and trigger project update"""
        # Call parent method
        self.ensure_one()
        if not self.pricing_id:
            raise UserError(_('Please select a pricing reference first!'))

        # Clear existing components
        self.component_line_ids.unlink()

        # Load components from pricing WITH SPECIFICATIONS
        component_lines = []
        for comp in self.pricing_id.component_line_ids:
            # Create component line
            comp_vals = {
                'component_id': comp.component_id.id,
                'quantity': comp.quantity,
                'weight': comp.weight,
                'cost_price': comp.cost_price,
                'bom_id': comp.bom_id.id if comp.bom_id else False,
            }
            component_lines.append((0, 0, comp_vals))

        self.write({
            'component_line_ids': component_lines,
            'state': 'components_loaded'
        })

        # Now copy specifications for each loaded component
        for planning_comp in self.component_line_ids:
            # Find the corresponding pricing component
            pricing_comp = self.pricing_id.component_line_ids.filtered(
                lambda c: c.component_id == planning_comp.component_id
            )

            if pricing_comp and pricing_comp.specification_ids:
                # Copy specifications
                for spec in pricing_comp.specification_ids:
                    self.env['component.specification.value'].create({
                        'planning_component_id': planning_comp.id,
                        'specification_id': spec.specification_id.id,
                        'value': spec.value,
                        'notes': spec.notes,
                        'sequence': spec.sequence,
                    })

        # Trigger project state update
        if self.project_id and self.project_id.auto_update_state:
            self.project_id.update_project_state()

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Success'),
                'message': _('Components and specifications loaded successfully!'),
                'type': 'success',
                'sticky': False,
            }
        }

    def action_material_planning(self):
        """Calculate material requirements and trigger project update"""
        self.ensure_one()
        if not self.component_line_ids:
            raise UserError(_('Please load components first!'))

        # Calculate material requirements
        self.material_requirement_ids.unlink()

        material_lines = []
        for comp in self.component_line_ids:
            if comp.bom_id:
                for bom_line in comp.bom_id.bom_line_ids:
                    required_qty = bom_line.product_qty * comp.quantity

                    product = bom_line.product_id
                    available_qty = product.qty_available - product.outgoing_qty

                    shortage_qty = max(0, required_qty - available_qty)

                    material_lines.append((0, 0, {
                        'component_id': comp.component_id.id,
                        'material_id': bom_line.product_id.id,
                        'required_qty': required_qty,
                        'available_qty': available_qty,
                        'shortage_qty': shortage_qty,
                    }))
            else:
                required_qty = comp.quantity
                product = comp.component_id
                available_qty = product.qty_available - product.outgoing_qty
                shortage_qty = max(0, required_qty - available_qty)

                material_lines.append((0, 0, {
                    'component_id': comp.component_id.id,
                    'material_id': comp.component_id.id,
                    'required_qty': required_qty,
                    'available_qty': available_qty,
                    'shortage_qty': shortage_qty,
                }))

        self.write({
            'material_requirement_ids': material_lines,
            'state': 'material_planned'
        })

        # Trigger project state update
        if self.project_id and self.project_id.auto_update_state:
            self.project_id.update_project_state()

        return {
            'name': _('Material Requirements'),
            'type': 'ir.actions.act_window',
            'res_model': 'material.requirement.wizard',
            'view_mode': 'form',
            'target': 'new',
            'context': {'default_planning_id': self.id}
        }

    def action_done(self):
        """Mark as done and update project"""
        self.write({'state': 'done'})

        # Trigger project state update
        for planning in self:
            if planning.project_id and planning.project_id.auto_update_state:
                planning.project_id.update_project_state()
    def _compute_production_count(self):
        for record in self:
            record.production_count = len(record.production_order_ids)

    @api.depends('production_order_ids.state', 'production_order_ids.product_qty', 'quantity')
    def _compute_produced_quantities(self):
        for record in self:
            main_productions = record.production_order_ids.filtered(
                lambda p: p.product_id == record.product_id
            )
            total_qty = sum(main_productions.mapped('product_qty'))
            record.total_produced_qty = total_qty
            record.remaining_qty = max(0, record.quantity - total_qty)

    @api.onchange('project_id')
    def _onchange_project_id(self):
        self.product_id = False
        self.pricing_id = False
        if self.project_id:
            product_ids = self.project_id.product_line_ids.mapped('product_id').ids
            return {'domain': {'product_id': [('id', 'in', product_ids)]}}
        return {'domain': {'product_id': []}}

    @api.onchange('product_id')
    def _onchange_product_id(self):
        self.pricing_id = False

        if self.product_id and self.project_id:
            return {
                'domain': {
                    'pricing_id': [
                        ('project_id', '=', self.project_id.id),
                        ('product_id', '=', self.product_id.id),
                        ('state', 'in', ['confirmed', 'approved'])
                    ]
                }
            }
        return {'domain': {'pricing_id': []}}

    def action_load_components(self):
        self.ensure_one()
        if not self.pricing_id:
            raise UserError(_('Please select a pricing reference first!'))

        # Clear existing components
        self.component_line_ids.unlink()

        # Load components from pricing WITH SPECIFICATIONS
        component_lines = []
        for comp in self.pricing_id.component_line_ids:
            # Create component line
            comp_vals = {
                'component_id': comp.component_id.id,
                'quantity': comp.quantity,
                'weight': comp.weight,
                'cost_price': comp.cost_price,
                'bom_id': comp.bom_id.id if comp.bom_id else False,
            }
            component_lines.append((0, 0, comp_vals))

        self.write({
            'component_line_ids': component_lines,
            'state': 'components_loaded'
        })

        # Now copy specifications for each loaded component
        for planning_comp in self.component_line_ids:
            # Find the corresponding pricing component
            pricing_comp = self.pricing_id.component_line_ids.filtered(
                lambda c: c.component_id == planning_comp.component_id
            )

            if pricing_comp and pricing_comp.specification_ids:
                # Copy specifications
                for spec in pricing_comp.specification_ids:
                    self.env['component.specification.value'].create({
                        'planning_component_id': planning_comp.id,
                        'specification_id': spec.specification_id.id,
                        'value': spec.value,
                        'notes': spec.notes,
                        'sequence': spec.sequence,
                    })

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Success'),
                'message': _('Components and specifications loaded successfully!'),
                'type': 'success',
                'sticky': False,
            }
        }

    def action_material_planning(self):
        self.ensure_one()
        if not self.component_line_ids:
            raise UserError(_('Please load components first!'))

        # Calculate material requirements
        self.material_requirement_ids.unlink()

        material_lines = []
        for comp in self.component_line_ids:
            if comp.bom_id:
                for bom_line in comp.bom_id.bom_line_ids:
                    required_qty = bom_line.product_qty * comp.quantity

                    product = bom_line.product_id
                    available_qty = product.qty_available - product.outgoing_qty

                    shortage_qty = max(0, required_qty - available_qty)

                    material_lines.append((0, 0, {
                        'component_id': comp.component_id.id,
                        'material_id': bom_line.product_id.id,
                        'required_qty': required_qty,
                        'available_qty': available_qty,
                        'shortage_qty': shortage_qty,
                    }))
            else:
                required_qty = comp.quantity
                product = comp.component_id
                available_qty = product.qty_available - product.outgoing_qty
                shortage_qty = max(0, required_qty - available_qty)

                material_lines.append((0, 0, {
                    'component_id': comp.component_id.id,
                    'material_id': comp.component_id.id,
                    'required_qty': required_qty,
                    'available_qty': available_qty,
                    'shortage_qty': shortage_qty,
                }))

        self.write({
            'material_requirement_ids': material_lines,
            'state': 'material_planned'
        })

        return {
            'name': _('Material Requirements'),
            'type': 'ir.actions.act_window',
            'res_model': 'material.requirement.wizard',
            'view_mode': 'form',
            'target': 'new',
            'context': {'default_planning_id': self.id}
        }

    def action_create_work_orders(self):
        self.ensure_one()

        if not self.component_line_ids:
            raise UserError(_('Please load components first!'))

        if self.remaining_qty <= 0:
            raise UserError(_(
                'Cannot create more work orders!\n'
                'Planned Quantity: %s\n'
                'Already Created: %s\n'
                'Remaining: %s'
            ) % (self.quantity, self.total_produced_qty, self.remaining_qty))

        wizard = self.env['work.order.creation.wizard'].create({
            'planning_id': self.id,
            'product_id': self.product_id.id,
            'max_quantity': self.remaining_qty,
            'quantity_to_produce': self.remaining_qty,
        })

        return {
            'name': _('Create Work Orders'),
            'type': 'ir.actions.act_window',
            'res_model': 'work.order.creation.wizard',
            'res_id': wizard.id,
            'view_mode': 'form',
            'target': 'new',
        }

    def action_done(self):
        self.write({'state': 'done'})

    def action_cancel(self):
        self.write({'state': 'cancelled'})

    def action_view_production_orders(self):
        """View all production orders created from this planning"""
        self.ensure_one()
        return {
            'name': _('Production Orders'),
            'type': 'ir.actions.act_window',
            'res_model': 'mrp.production',
            'view_mode': 'tree,form',
            'domain': [('id', 'in', self.production_order_ids.ids)],
            'context': {'default_origin': self.name},
        }

    def action_sync_specifications_from_pricing(self):
        """Manually sync specifications from pricing (in case pricing was updated)"""
        self.ensure_one()

        if not self.pricing_id:
            raise UserError(_('No pricing reference selected!'))

        synced_count = 0

        for planning_comp in self.component_line_ids:
            # Find corresponding pricing component
            pricing_comp = self.pricing_id.component_line_ids.filtered(
                lambda c: c.component_id == planning_comp.component_id
            )

            if pricing_comp:
                # Delete existing specifications
                planning_comp.specification_ids.unlink()

                # Copy new specifications
                for spec in pricing_comp.specification_ids:
                    self.env['component.specification.value'].create({
                        'planning_component_id': planning_comp.id,
                        'specification_id': spec.specification_id.id,
                        'value': spec.value,
                        'notes': spec.notes,
                        'sequence': spec.sequence,
                    })
                    synced_count += 1

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Success'),
                'message': _('%s specifications synced from pricing!') % synced_count,
                'type': 'success',
                'sticky': False,
            }
        }


class MaterialPlanningComponent(models.Model):
    _name = 'material.planning.component'
    _description = 'Material Planning Component'
    _order = 'sequence, id'

    sequence = fields.Integer(string='Sequence', default=10)
    planning_id = fields.Many2one(
        'material.production.planning',
        string='Planning',
        required=True,
        ondelete='cascade'
    )
    component_id = fields.Many2one(
        'product.product',
        string='Component Product',
        required=True
    )
    quantity = fields.Float(
        string='Quantity',
        required=True,
        digits='Product Unit of Measure'
    )
    weight = fields.Float(
        string='Weight',
        digits='Stock Weight'
    )
    cost_price = fields.Float(
        string='Cost Price',
        digits='Product Price'
    )
    bom_id = fields.Many2one(
        'mrp.bom',
        string='Bill of Materials'
    )
    specification_ids = fields.One2many(
        'component.specification.value',
        'planning_component_id',
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

    uom_id = fields.Many2one(
        'uom.uom',
        string='Unit of Measure',
        related='component_id.uom_id',
        readonly=True
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

    def action_component_specifications(self):
        """Open specifications wizard (read-only view since source is pricing)"""
        self.ensure_one()
        return {
            'type': 'ir.actions.act_window',
            'name': _('Component Specifications: %s') % self.component_id.name,
            'res_model': 'component.specification.value',
            'view_mode': 'tree,form',
            'domain': [('planning_component_id', '=', self.id)],
            'context': {
                'default_planning_component_id': self.id,
            },
            'target': 'current',
        }


class MaterialRequirementLine(models.Model):
    _name = 'material.requirement.line'
    _description = 'Material Requirement Line'

    planning_id = fields.Many2one(
        'material.production.planning',
        string='Planning',
        required=True,
        ondelete='cascade'
    )
    component_id = fields.Many2one(
        'product.product',
        string='For Component'
    )
    material_id = fields.Many2one(
        'product.product',
        string='Raw Material',
        required=True
    )
    required_qty = fields.Float(
        string='Required Quantity',
        digits='Product Unit of Measure'
    )
    available_qty = fields.Float(
        string='Available Stock',
        digits='Product Unit of Measure'
    )
    shortage_qty = fields.Float(
        string='Shortage',
        digits='Product Unit of Measure'
    )
    uom_id = fields.Many2one(
        'uom.uom',
        string='Unit of Measure',
        related='material_id.uom_id',
        readonly=True
    )