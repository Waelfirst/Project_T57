# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import UserError


class ComponentSpecificationDefinition(models.Model):
    """تعريف المواصفات - Specification Definition"""
    _name = 'component.specification.definition'
    _description = 'Component Specification Definition'
    _order = 'sequence, id'

    name = fields.Char(
        string='Specification Name',
        required=True,
        translate=True,
        help='اسم المواصفة (مثال: المادة، اللون، الأبعاد)'
    )
    code = fields.Char(
        string='Code',
        help='كود فريد للمواصفة'
    )
    sequence = fields.Integer(string='Sequence', default=10)
    active = fields.Boolean(string='Active', default=True)
    description = fields.Text(string='Description')


class ComponentSpecificationValue(models.Model):
    """قيم المواصفات للأجزاء - Specification Values for Components"""
    _name = 'component.specification.value'
    _description = 'Component Specification Value'
    _order = 'sequence, id'

    sequence = fields.Integer(string='Sequence', default=10)
    
    # Links to parent records
    pricing_component_id = fields.Many2one(
        'project.product.component',
        string='Pricing Component',
        ondelete='cascade'
    )
    planning_component_id = fields.Many2one(
        'material.planning.component',
        string='Planning Component',
        ondelete='cascade'
    )
    
    # Specification details
    specification_id = fields.Many2one(
        'component.specification.definition',
        string='Specification',
        required=True
    )
    specification_name = fields.Char(
        related='specification_id.name',
        string='Specification Name',
        store=True
    )
    value = fields.Char(
        string='Value',
        required=True,
        translate=True,
        help='قيمة المواصفة'
    )
    notes = fields.Text(string='Notes')


class ComponentSpecificationWizard(models.TransientModel):
    """معالج إدخال المواصفات - Specification Entry Wizard"""
    _name = 'component.specification.wizard'
    _description = 'Component Specification Wizard'

    component_id = fields.Many2one(
        'product.product',
        string='Component',
        required=True,
        readonly=True
    )
    component_name = fields.Char(
        related='component_id.name',
        string='Component Name',
        readonly=True
    )
    
    # Context fields
    source_model = fields.Char(string='Source Model')
    source_id = fields.Integer(string='Source Record ID')
    
    specification_line_ids = fields.One2many(
        'component.specification.wizard.line',
        'wizard_id',
        string='Specifications'
    )
    
    @api.model
    def default_get(self, fields_list):
        res = super(ComponentSpecificationWizard, self).default_get(fields_list)
        
        context = self.env.context
        source_model = context.get('source_model')
        source_id = context.get('source_id')
        component_id = context.get('component_id')
        
        if source_model and source_id and component_id:
            res['source_model'] = source_model
            res['source_id'] = source_id
            res['component_id'] = component_id
            
            # Load existing specifications
            spec_lines = []
            if source_model == 'project.product.component':
                existing_specs = self.env['component.specification.value'].search([
                    ('pricing_component_id', '=', source_id)
                ])
            elif source_model == 'material.planning.component':
                existing_specs = self.env['component.specification.value'].search([
                    ('planning_component_id', '=', source_id)
                ])
            else:
                existing_specs = self.env['component.specification.value']
            
            # Add existing specs to wizard
            for spec in existing_specs:
                spec_lines.append((0, 0, {
                    'specification_id': spec.specification_id.id,
                    'value': spec.value,
                    'notes': spec.notes,
                    'sequence': spec.sequence,
                }))
            
            res['specification_line_ids'] = spec_lines
        
        return res
    
    def action_save_specifications(self):
        self.ensure_one()
        
        # Delete existing specifications
        if self.source_model == 'project.product.component':
            self.env['component.specification.value'].search([
                ('pricing_component_id', '=', self.source_id)
            ]).unlink()
        elif self.source_model == 'material.planning.component':
            self.env['component.specification.value'].search([
                ('planning_component_id', '=', self.source_id)
            ]).unlink()
        
        # Create new specifications
        for line in self.specification_line_ids:
            if line.specification_id and line.value:  # Only save if both are filled
                vals = {
                    'specification_id': line.specification_id.id,
                    'value': line.value,
                    'notes': line.notes,
                    'sequence': line.sequence,
                }
                
                if self.source_model == 'project.product.component':
                    vals['pricing_component_id'] = self.source_id
                elif self.source_model == 'material.planning.component':
                    vals['planning_component_id'] = self.source_id
                
                self.env['component.specification.value'].create(vals)
        
        return {'type': 'ir.actions.act_window_close'}


class ComponentSpecificationWizardLine(models.TransientModel):
    _name = 'component.specification.wizard.line'
    _description = 'Component Specification Wizard Line'
    _order = 'sequence, id'

    wizard_id = fields.Many2one(
        'component.specification.wizard',
        string='Wizard',
        required=True,
        ondelete='cascade'
    )
    sequence = fields.Integer(string='Sequence', default=10)
    specification_id = fields.Many2one(
        'component.specification.definition',
        string='Specification',
        required=True
    )
    value = fields.Char(
        string='Value',
        required=True,
        translate=True
    )
    notes = fields.Text(string='Notes')
