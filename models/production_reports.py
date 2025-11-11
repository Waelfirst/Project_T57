# -*- coding: utf-8 -*-

from odoo import models, fields, api, _


class ProductionProgressReport(models.Model):
    _name = 'production.progress.report'
    _description = 'Production Progress Report'
    _auto = False
    _order = 'project_id, product_id, component_id'

    project_id = fields.Many2one('project.definition', string='Project', readonly=True)
    product_id = fields.Many2one('product.product', string='Product', readonly=True)
    component_id = fields.Many2one('product.product', string='Component', readonly=True)
    
    # Pricing data
    planned_quantity = fields.Float(string='Planned Quantity', readonly=True)
    planned_weight = fields.Float(string='Planned Weight', readonly=True)
    
    # Production data
    produced_quantity = fields.Float(string='Produced Quantity', readonly=True)
    production_state = fields.Char(string='Production State', readonly=True)
    current_operation = fields.Char(string='Current Operation', readonly=True)
    progress_percentage = fields.Float(string='Progress %', readonly=True)
    
    # Status
    status = fields.Selection([
        ('not_started', 'Not Started'),
        ('in_progress', 'In Progress'),
        ('completed', 'Completed'),
        ('cancelled', 'Cancelled'),
    ], string='Status', readonly=True)

    def init(self):
        self._cr.execute("""
            CREATE OR REPLACE VIEW production_progress_report AS (
                SELECT
                    ROW_NUMBER() OVER() as id,
                    pp.project_id,
                    pp.product_id,
                    ppc.component_id,
                    ppc.quantity as planned_quantity,
                    ppc.weight as planned_weight,
                    COALESCE(mp.product_qty, 0) as produced_quantity,
                    mp.state as production_state,
                    CASE 
                        WHEN mp.state IS NULL THEN 'not_started'
                        WHEN mp.state = 'done' THEN 'completed'
                        WHEN mp.state = 'cancel' THEN 'cancelled'
                        ELSE 'in_progress'
                    END as status,
                    COALESCE(
                        (SELECT mw.name 
                         FROM mrp_workorder mw 
                         WHERE mw.production_id = mp.id 
                         AND mw.state IN ('ready', 'progress')
                         LIMIT 1),
                        'N/A'
                    ) as current_operation,
                    CASE 
                        WHEN mp.id IS NOT NULL THEN
                            COALESCE(
                                (SELECT COUNT(*)::float / NULLIF(COUNT(*), 0) * 100
                                 FROM mrp_workorder mw1
                                 WHERE mw1.production_id = mp.id
                                 AND mw1.state = 'done'),
                                0
                            )
                        ELSE 0
                    END as progress_percentage
                FROM 
                    project_product_pricing pp
                JOIN 
                    project_product_component ppc ON ppc.pricing_id = pp.id
                LEFT JOIN 
                    mrp_production mp ON mp.product_id = ppc.component_id
                    AND mp.origin LIKE '%' || (SELECT name FROM project_definition WHERE id = pp.project_id) || '%'
                WHERE 
                    pp.state IN ('confirmed', 'approved')
            )
        """)


class MaterialUsageReport(models.Model):
    _name = 'material.usage.report'
    _description = 'Material Usage Report'
    _auto = False
    _order = 'project_id, product_id, material_id'

    project_id = fields.Many2one('project.definition', string='Project', readonly=True)
    product_id = fields.Many2one('product.product', string='Product', readonly=True)
    component_id = fields.Many2one('product.product', string='Component', readonly=True)
    material_id = fields.Many2one('product.product', string='Material', readonly=True)
    
    # Requirements
    required_quantity = fields.Float(string='Required Quantity', readonly=True)
    
    # Stock
    available_quantity = fields.Float(string='Available Quantity', readonly=True)
    reserved_quantity = fields.Float(string='Reserved Quantity', readonly=True)
    
    # Consumption
    consumed_quantity = fields.Float(string='Consumed Quantity', readonly=True)
    
    # Purchasing
    ordered_quantity = fields.Float(string='Ordered Quantity', readonly=True)
    received_quantity = fields.Float(string='Received Quantity', readonly=True)
    
    # Status
    shortage_quantity = fields.Float(string='Shortage', readonly=True)
    status = fields.Selection([
        ('sufficient', 'Sufficient'),
        ('partial', 'Partially Available'),
        ('shortage', 'Shortage'),
        ('ordered', 'Ordered'),
        ('received', 'Received'),
    ], string='Status', readonly=True)

    def init(self):
        self._cr.execute("""
            CREATE OR REPLACE VIEW material_usage_report AS (
                SELECT
                    ROW_NUMBER() OVER() as id,
                    pp.project_id,
                    pp.product_id,
                    ppc.component_id,
                    mbl.product_id as material_id,
                    
                    -- Required quantity
                    mbl.product_qty * ppc.quantity as required_quantity,
                    
                    -- Stock availability
                    COALESCE(sq.quantity, 0) as available_quantity,
                    COALESCE(sq.reserved_quantity, 0) as reserved_quantity,
                    
                    -- Consumed
                    COALESCE(
                        (SELECT SUM(sm.product_uom_qty)
                         FROM stock_move sm
                         WHERE sm.product_id = mbl.product_id
                         AND sm.raw_material_production_id IN (
                             SELECT id FROM mrp_production mp2
                             WHERE mp2.product_id = ppc.component_id
                             AND mp2.origin LIKE '%' || (SELECT name FROM project_definition WHERE id = pp.project_id) || '%'
                         )
                         AND sm.state = 'done'),
                        0
                    ) as consumed_quantity,
                    
                    -- Purchasing
                    COALESCE(
                        (SELECT SUM(pol.product_qty)
                         FROM purchase_order_line pol
                         JOIN purchase_order po ON po.id = pol.order_id
                         WHERE pol.product_id = mbl.product_id
                         AND po.origin LIKE '%' || (SELECT name FROM project_definition WHERE id = pp.project_id) || '%'
                         AND po.state IN ('draft', 'sent', 'to approve', 'purchase')),
                        0
                    ) as ordered_quantity,
                    
                    COALESCE(
                        (SELECT SUM(pol.qty_received)
                         FROM purchase_order_line pol
                         JOIN purchase_order po ON po.id = pol.order_id
                         WHERE pol.product_id = mbl.product_id
                         AND po.origin LIKE '%' || (SELECT name FROM project_definition WHERE id = pp.project_id) || '%'
                         AND po.state IN ('purchase', 'done')),
                        0
                    ) as received_quantity,
                    
                    -- Shortage
                    GREATEST(
                        (mbl.product_qty * ppc.quantity) - COALESCE(sq.quantity, 0) - COALESCE(sq.reserved_quantity, 0),
                        0
                    ) as shortage_quantity,
                    
                    -- Status
                    CASE 
                        WHEN COALESCE(sq.quantity, 0) >= (mbl.product_qty * ppc.quantity) THEN 'sufficient'
                        WHEN COALESCE(
                            (SELECT SUM(pol.qty_received)
                             FROM purchase_order_line pol
                             JOIN purchase_order po ON po.id = pol.order_id
                             WHERE pol.product_id = mbl.product_id
                             AND po.origin LIKE '%' || (SELECT name FROM project_definition WHERE id = pp.project_id) || '%'),
                            0
                        ) > 0 THEN 'received'
                        WHEN COALESCE(
                            (SELECT SUM(pol.product_qty)
                             FROM purchase_order_line pol
                             JOIN purchase_order po ON po.id = pol.order_id
                             WHERE pol.product_id = mbl.product_id
                             AND po.origin LIKE '%' || (SELECT name FROM project_definition WHERE id = pp.project_id) || '%'),
                            0
                        ) > 0 THEN 'ordered'
                        WHEN COALESCE(sq.quantity, 0) > 0 THEN 'partial'
                        ELSE 'shortage'
                    END as status
                    
                FROM 
                    project_product_pricing pp
                JOIN 
                    project_product_component ppc ON ppc.pricing_id = pp.id
                JOIN 
                    mrp_bom mb ON mb.id = ppc.bom_id
                JOIN 
                    mrp_bom_line mbl ON mbl.bom_id = mb.id
                LEFT JOIN 
                    stock_quant sq ON sq.product_id = mbl.product_id
                    AND sq.location_id IN (SELECT id FROM stock_location WHERE usage = 'internal')
                WHERE 
                    pp.state IN ('confirmed', 'approved')
                    AND ppc.bom_id IS NOT NULL
            )
        """)
