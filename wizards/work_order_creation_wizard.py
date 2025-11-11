# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import UserError, ValidationError


class WorkOrderCreationWizard(models.TransientModel):
    _name = 'work.order.creation.wizard'
    _description = 'Work Order Creation Wizard'

    planning_id = fields.Many2one(
        'material.production.planning',
        string='Planning',
        required=True,
        readonly=True
    )
    product_id = fields.Many2one(
        'product.product',
        string='Product',
        required=True,
        readonly=True
    )
    max_quantity = fields.Float(
        string='Maximum Quantity',
        readonly=True,
        help='Maximum quantity that can be produced'
    )
    quantity_to_produce = fields.Float(
        string='Quantity to Produce',
        required=True,
        default=0.0
    )
    
    create_component_orders = fields.Boolean(
        string='Create Component Work Orders',
        default=True,
        help='Also create work orders for components'
    )
    
    @api.onchange('quantity_to_produce', 'planning_id', 'create_component_orders')
    def _onchange_quantity_preview(self):
        """Update component preview when quantity changes"""
        if self.planning_id and self.quantity_to_produce > 0:
            ratio = self.quantity_to_produce / self.planning_id.quantity if self.planning_id.quantity > 0 else 1
            
            preview_text = _("Component Orders to Create:\n\n")
            for comp in self.planning_id.component_line_ids:
                if comp.bom_id:
                    component_qty = comp.quantity * ratio
                    preview_text += _("• %s: %.2f units (BOM: %s)\n") % (
                        comp.component_id.name,
                        component_qty,
                        comp.bom_id.code or comp.bom_id.id
                    )
            
            self.component_preview = preview_text if preview_text != _("Component Orders to Create:\n\n") else _("No components with BOM found")
    
    component_preview = fields.Text(
        string='Component Preview',
        readonly=True,
        compute='_compute_component_preview',
        store=False
    )
    
    
    @api.depends('quantity_to_produce', 'planning_id')
    def _compute_component_preview(self):
        for wizard in self:
            if wizard.planning_id and wizard.quantity_to_produce > 0:
                ratio = wizard.quantity_to_produce / wizard.planning_id.quantity if wizard.planning_id.quantity > 0 else 1
                
                preview_text = _("Component Orders to Create:\n\n")
                for comp in wizard.planning_id.component_line_ids:
                    if comp.bom_id:
                        component_qty = comp.quantity * ratio
                        preview_text += _("• %s: %.2f units (BOM: %s)\n") % (
                            comp.component_id.name,
                            component_qty,
                            comp.bom_id.code or comp.bom_id.id
                        )
                
                wizard.component_preview = preview_text if preview_text != _("Component Orders to Create:\n\n") else _("No components with BOM found")
            else:
                wizard.component_preview = _("Enter quantity to see component preview")
    
    @api.constrains('quantity_to_produce')
    def _check_quantity(self):
        for wizard in self:
            if wizard.quantity_to_produce <= 0:
                raise ValidationError(_('Quantity to produce must be greater than zero!'))
            
            if wizard.quantity_to_produce > wizard.max_quantity:
                raise ValidationError(_(
                    'Quantity to produce (%s) cannot exceed remaining quantity (%s)!'
                ) % (wizard.quantity_to_produce, wizard.max_quantity))
    
    def action_create_orders(self):
        self.ensure_one()
        
        # Validate quantity
        if self.quantity_to_produce > self.max_quantity:
            raise UserError(_(
                'Cannot produce %s units!\n'
                'Maximum allowed: %s'
            ) % (self.quantity_to_produce, self.max_quantity))
        
        production_ids = []
        
        # Create main product production order
        main_production = self.env['mrp.production'].create({
            'product_id': self.product_id.id,
            'product_qty': self.quantity_to_produce,
            'product_uom_id': self.product_id.uom_id.id,
            'origin': self.planning_id.name,
        })
        production_ids.append(main_production.id)
        
        # Create component production orders if requested
        if self.create_component_orders:
            ratio = self.quantity_to_produce / self.planning_id.quantity if self.planning_id.quantity > 0 else 1
            
            for comp in self.planning_id.component_line_ids:
                if comp.bom_id:
                    # Calculate component quantity based on ratio
                    component_qty = comp.quantity * ratio
                    
                    # Validate component quantity against what's already been produced
                    existing_productions = self.env['mrp.production'].search([
                        ('origin', 'like', self.planning_id.name),
                        ('product_id', '=', comp.component_id.id),
                    ])
                    total_existing = sum(existing_productions.mapped('product_qty'))
                    
                    # Check if we're exceeding the planned quantity for this component
                    if total_existing + component_qty > comp.quantity:
                        raise UserError(_(
                            'Cannot create work order for %s!\n'
                            'Planned: %s\n'
                            'Already created: %s\n'
                            'Trying to create: %s\n'
                            'This would exceed planned quantity!'
                        ) % (comp.component_id.name, comp.quantity, total_existing, component_qty))
                    
                    comp_production = self.env['mrp.production'].create({
                        'product_id': comp.component_id.id,
                        'product_qty': component_qty,
                        'product_uom_id': comp.component_id.uom_id.id,
                        'bom_id': comp.bom_id.id,
                        'origin': f"{self.planning_id.name} - {comp.component_id.name}",
                    })
                    production_ids.append(comp_production.id)
        
        # Link productions to planning
        self.planning_id.write({
            'production_order_ids': [(4, pid) for pid in production_ids],
            'state': 'work_orders_created'
        })
        
        # Return action to view created production orders
        return {
            'type': 'ir.actions.act_window',
            'name': _('Production Orders'),
            'res_model': 'mrp.production',
            'view_mode': 'tree,form',
            'domain': [('id', 'in', production_ids)],
            'target': 'current',
        }
