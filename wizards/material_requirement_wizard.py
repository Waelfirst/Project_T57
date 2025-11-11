# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import UserError


class MaterialRequirementWizard(models.TransientModel):
    _name = 'material.requirement.wizard'
    _description = 'Material Requirement Wizard'

    planning_id = fields.Many2one(
        'material.production.planning',
        string='Planning',
        required=True
    )
    
    material_line_ids = fields.One2many(
        related='planning_id.material_requirement_ids',
        string='Material Requirements',
        readonly=True
    )
    
    def action_create_rfq(self):
        self.ensure_one()
        
        # Get materials with shortage
        shortage_lines = self.planning_id.material_requirement_ids.filtered(
            lambda l: l.shortage_qty > 0
        )
        
        if not shortage_lines:
            raise UserError(_('No material shortage found!'))
        
        # Group by product to avoid duplicate lines
        product_qty_map = {}
        for line in shortage_lines:
            if line.material_id.id in product_qty_map:
                product_qty_map[line.material_id.id] += line.shortage_qty
            else:
                product_qty_map[line.material_id.id] = line.shortage_qty
        
        # Create RFQ
        po_lines = []
        for product_id, qty in product_qty_map.items():
            product = self.env['product.product'].browse(product_id)
            
            # Find supplier
            supplier_info = product.seller_ids[:1]
            
            po_lines.append((0, 0, {
                'product_id': product_id,
                'product_qty': qty,
                'product_uom': product.uom_po_id.id,
                'price_unit': supplier_info.price if supplier_info else product.standard_price,
                'name': product.name,
                'date_planned': fields.Datetime.now(),
            }))
        
        # Get default supplier or create without supplier
        default_supplier = False
        if shortage_lines and shortage_lines[0].material_id.seller_ids:
            default_supplier = shortage_lines[0].material_id.seller_ids[0].partner_id.id
        
        purchase_order = self.env['purchase.order'].create({
            'partner_id': default_supplier if default_supplier else self.env.ref('base.main_partner').id,
            'origin': self.planning_id.name,
            'order_line': po_lines,
        })
        
        # Link RFQ to planning
        self.planning_id.rfq_ids = [(4, purchase_order.id)]
        
        return {
            'name': _('Request for Quotation'),
            'type': 'ir.actions.act_window',
            'res_model': 'purchase.order',
            'view_mode': 'form',
            'res_id': purchase_order.id,
            'target': 'current',
        }
