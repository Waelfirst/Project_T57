# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import UserError


class UserPermissionWizard(models.TransientModel):
    _name = 'user.permission.wizard'
    _description = 'User Permission Assignment Wizard'

    user_id = fields.Many2one(
        'res.users',
        string='User',
        required=True,
        help='Select user to assign permissions'
    )

    template_type = fields.Selection([
        ('full', 'Full Access - All Functions'),
        ('read_only', 'Read Only - View Only'),
        ('custom', 'Custom - Configure Later'),
    ], string='Permission Template', default='full', required=True)

    screen_ids = fields.Many2many(
        'screen.definition',
        string='Screens',
        help='Select screens to grant access'
    )

    def action_assign_permissions(self):
        """Assign permissions to the selected user"""
        self.ensure_one()

        if not self.screen_ids:
            raise UserError(_('Please select at least one screen!'))

        created_count = 0
        updated_count = 0

        for screen in self.screen_ids:
            # Check if permission already exists
            existing_permission = self.env['user.screen.permission'].search([
                ('user_id', '=', self.user_id.id),
                ('screen_id', '=', screen.id),
            ], limit=1)

            if existing_permission:
                # Update existing permission
                existing_permission.write({
                    'has_access': True,
                    'access_level': self.template_type,
                })
                updated_count += 1
            else:
                # Create new permission
                self.env['user.screen.permission'].create({
                    'user_id': self.user_id.id,
                    'screen_id': screen.id,
                    'has_access': True,
                    'access_level': self.template_type,
                })
                created_count += 1

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Success'),
                'message': _('Permissions assigned successfully!\n'
                             'Created: %s\nUpdated: %s') % (created_count, updated_count),
                'type': 'success',
                'sticky': False,
            }
        }