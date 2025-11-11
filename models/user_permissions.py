# -*- coding: utf-8 -*-

from odoo import models, fields, api, _
from odoo.exceptions import UserError, AccessError
import logging

_logger = logging.getLogger(__name__)


class ScreenDefinition(models.Model):
    """تعريف الشاشات في النظام - Screen Definitions"""
    _name = 'screen.definition'
    _description = 'Screen Definition'
    _order = 'sequence, name'

    name = fields.Char(string='Screen Name', required=True, translate=True)
    code = fields.Char(string='Technical Code', required=True, help='Technical identifier for the screen')
    model_name = fields.Char(string='Model Name', required=True, help='Odoo model name (e.g., project.definition)')
    sequence = fields.Integer(string='Sequence', default=10)
    active = fields.Boolean(string='Active', default=True)
    description = fields.Text(string='Description')

    function_ids = fields.One2many(
        'screen.function.definition',
        'screen_id',
        string='Available Functions'
    )

    user_permission_ids = fields.One2many(
        'user.screen.permission',
        'screen_id',
        string='User Permissions'
    )

    _sql_constraints = [
        ('code_unique', 'UNIQUE(code)', 'Screen code must be unique!'),
        ('model_name_unique', 'UNIQUE(model_name)', 'Model name must be unique!'),
    ]


class ScreenFunctionDefinition(models.Model):
    """تعريف الوظائف المتاحة في كل شاشة - Screen Functions"""
    _name = 'screen.function.definition'
    _description = 'Screen Function Definition'
    _order = 'sequence, name'

    name = fields.Char(string='Function Name', required=True, translate=True)
    code = fields.Char(string='Technical Code', required=True, help='Method name or action identifier')
    screen_id = fields.Many2one('screen.definition', string='Screen', required=True, ondelete='cascade')
    function_type = fields.Selection([
        ('create', 'Create/Add'),
        ('write', 'Edit/Update'),
        ('read', 'View/Read'),
        ('unlink', 'Delete'),
        ('import', 'Import'),
        ('export', 'Export'),
        ('action', 'Custom Action'),
        ('report', 'Report'),
    ], string='Function Type', required=True, default='action')
    sequence = fields.Integer(string='Sequence', default=10)
    active = fields.Boolean(string='Active', default=True)
    description = fields.Text(string='Description')

    _sql_constraints = [
        ('code_screen_unique', 'UNIQUE(code, screen_id)', 'Function code must be unique per screen!'),
    ]


class UserScreenPermission(models.Model):
    """صلاحيات المستخدمين على الشاشات - User Screen Permissions"""
    _name = 'user.screen.permission'
    _description = 'User Screen Permission'
    _order = 'user_id, screen_id'

    user_id = fields.Many2one('res.users', string='User', required=True, ondelete='cascade')
    screen_id = fields.Many2one('screen.definition', string='Screen', required=True, ondelete='cascade')

    # Overall screen access
    has_access = fields.Boolean(string='Has Access', default=True)
    access_level = fields.Selection([
        ('full', 'Full Access'),
        ('read_only', 'Read Only'),
        ('custom', 'Custom Functions'),
        ('no_access', 'No Access'),
    ], string='Access Level', default='custom', required=True)

    # Function-specific permissions
    function_permission_ids = fields.One2many(
        'user.function.permission',
        'user_screen_permission_id',
        string='Function Permissions'
    )

    # Quick access fields
    can_create = fields.Boolean(string='Can Create', compute='_compute_quick_access', store=True)
    can_read = fields.Boolean(string='Can Read', compute='_compute_quick_access', store=True)
    can_write = fields.Boolean(string='Can Write', compute='_compute_quick_access', store=True)
    can_delete = fields.Boolean(string='Can Delete', compute='_compute_quick_access', store=True)
    can_import = fields.Boolean(string='Can Import', compute='_compute_quick_access', store=True)
    can_export = fields.Boolean(string='Can Export', compute='_compute_quick_access', store=True)

    notes = fields.Text(string='Notes')

    _sql_constraints = [
        ('user_screen_unique', 'UNIQUE(user_id, screen_id)', 'User can have only one permission record per screen!'),
    ]

    @api.depends('access_level', 'function_permission_ids', 'function_permission_ids.permission_type')
    def _compute_quick_access(self):
        for record in self:
            if record.access_level == 'full':
                record.can_create = True
                record.can_read = True
                record.can_write = True
                record.can_delete = True
                record.can_import = True
                record.can_export = True
            elif record.access_level == 'read_only':
                record.can_create = False
                record.can_read = True
                record.can_write = False
                record.can_delete = False
                record.can_import = False
                record.can_export = True
            elif record.access_level == 'no_access':
                record.can_create = False
                record.can_read = False
                record.can_write = False
                record.can_delete = False
                record.can_import = False
                record.can_export = False
            else:  # custom
                # Calculate based on function permissions
                func_perms = record.function_permission_ids
                record.can_create = any(
                    f.function_id.function_type == 'create' and f.permission_type in ['full', 'execute'] for f in
                    func_perms)
                record.can_read = any(
                    f.function_id.function_type == 'read' and f.permission_type in ['full', 'execute', 'read_only'] for
                    f in func_perms)
                record.can_write = any(
                    f.function_id.function_type == 'write' and f.permission_type in ['full', 'execute'] for f in
                    func_perms)
                record.can_delete = any(
                    f.function_id.function_type == 'unlink' and f.permission_type in ['full', 'execute'] for f in
                    func_perms)
                record.can_import = any(
                    f.function_id.function_type == 'import' and f.permission_type in ['full', 'execute'] for f in
                    func_perms)
                record.can_export = any(
                    f.function_id.function_type == 'export' and f.permission_type in ['full', 'execute'] for f in
                    func_perms)

    @api.onchange('screen_id')
    def _onchange_screen_id(self):
        """Load available functions when screen is selected"""
        if self.screen_id:
            # Get existing function permissions
            existing_funcs = self.function_permission_ids.mapped('function_id')

            # Add new functions
            new_funcs = []
            for func in self.screen_id.function_ids:
                if func not in existing_funcs:
                    new_funcs.append((0, 0, {
                        'function_id': func.id,
                        'permission_type': 'no_access',
                    }))

            if new_funcs:
                self.function_permission_ids = new_funcs

    @api.model
    def create(self, vals):
        """Auto-create function permissions when screen permission is created"""
        record = super(UserScreenPermission, self).create(vals)

        # Create function permissions for all available functions
        if record.screen_id and not record.function_permission_ids:
            func_perms = []
            for func in record.screen_id.function_ids:
                func_perms.append((0, 0, {
                    'function_id': func.id,
                    'permission_type': 'no_access' if record.access_level == 'custom' else 'full',
                }))

            if func_perms:
                record.function_permission_ids = func_perms

        return record


class UserFunctionPermission(models.Model):
    """صلاحيات المستخدمين على الوظائف - User Function Permissions"""
    _name = 'user.function.permission'
    _description = 'User Function Permission'
    _order = 'function_id'

    user_screen_permission_id = fields.Many2one(
        'user.screen.permission',
        string='User Screen Permission',
        required=True,
        ondelete='cascade'
    )
    user_id = fields.Many2one(related='user_screen_permission_id.user_id', string='User', store=True, readonly=True)
    screen_id = fields.Many2one(related='user_screen_permission_id.screen_id', string='Screen', store=True,
                                readonly=True)

    function_id = fields.Many2one(
        'screen.function.definition',
        string='Function',
        required=True,
        domain="[('screen_id', '=', screen_id)]"
    )
    function_name = fields.Char(related='function_id.name', string='Function Name', readonly=True)
    function_type = fields.Selection(related='function_id.function_type', string='Type', readonly=True)

    permission_type = fields.Selection([
        ('full', 'Full Access - إدخال كامل'),
        ('execute', 'Execute - تنفيذ فقط'),
        ('read_only', 'Read Only - استعلام فقط'),
        ('no_access', 'No Access - إلغاء'),
    ], string='Permission Type', required=True, default='no_access')

    notes = fields.Text(string='Notes')


class PermissionHelper(models.AbstractModel):
    """Helper model for permission checks"""
    _name = 'permission.helper'
    _description = 'Permission Helper'

    @api.model
    def check_user_permission(self, screen_code, function_code=None, permission_required='execute'):
        """
        Check if current user has permission for a screen/function

        Args:
            screen_code: Technical code of the screen
            function_code: Technical code of the function (optional)
            permission_required: Type of permission needed ('execute', 'read_only', 'full')

        Returns:
            Boolean: True if user has permission, False otherwise
        """
        user = self.env.user

        # Superuser always has access
        if user.id == self.env.ref('base.user_admin').id:
            return True

        # Find screen
        screen = self.env['screen.definition'].search([('code', '=', screen_code)], limit=1)
        if not screen:
            _logger.warning('Screen not found: %s', screen_code)
            return False

        # Find user permission
        user_permission = self.env['user.screen.permission'].search([
            ('user_id', '=', user.id),
            ('screen_id', '=', screen.id),
        ], limit=1)

        if not user_permission:
            _logger.warning('No permission record found for user %s on screen %s', user.name, screen_code)
            return False

        # Check access level
        if not user_permission.has_access or user_permission.access_level == 'no_access':
            return False

        if user_permission.access_level == 'full':
            return True

        if user_permission.access_level == 'read_only' and permission_required in ['read_only', 'execute']:
            return True

        # Check function-specific permission
        if function_code:
            function = self.env['screen.function.definition'].search([
                ('screen_id', '=', screen.id),
                ('code', '=', function_code),
            ], limit=1)

            if not function:
                _logger.warning('Function not found: %s in screen %s', function_code, screen_code)
                return False

            func_permission = self.env['user.function.permission'].search([
                ('user_screen_permission_id', '=', user_permission.id),
                ('function_id', '=', function.id),
            ], limit=1)

            if not func_permission:
                return False

            # Check permission type
            if permission_required == 'full':
                return func_permission.permission_type == 'full'
            elif permission_required == 'execute':
                return func_permission.permission_type in ['full', 'execute']
            elif permission_required == 'read_only':
                return func_permission.permission_type in ['full', 'execute', 'read_only']

        return False

    @api.model
    def raise_permission_error(self, screen_code, function_code=None):
        """Raise user-friendly permission error"""
        screen = self.env['screen.definition'].search([('code', '=', screen_code)], limit=1)
        screen_name = screen.name if screen else screen_code

        if function_code:
            function = self.env['screen.function.definition'].search([
                ('screen_id', '=', screen.id),
                ('code', '=', function_code),
            ], limit=1) if screen else False
            function_name = function.name if function else function_code

            raise AccessError(_(
                '⛔ Access Denied!\n\n'
                'You do not have permission to execute "%s" in "%s".\n\n'
                'Please contact your system administrator to request access.'
            ) % (function_name, screen_name))
        else:
            raise AccessError(_(
                '⛔ Access Denied!\n\n'
                'You do not have permission to access "%s".\n\n'
                'Please contact your system administrator to request access.'
            ) % screen_name)


def check_permission(screen_code, function_code=None, permission_type='execute'):
    """
    Decorator to check user permissions before executing a method

    Usage:
        @check_permission('project_definition', 'action_confirm', 'execute')
        def action_confirm(self):
            ...
    """

    def decorator(func):
        def wrapper(self, *args, **kwargs):
            # Check permission
            helper = self.env['permission.helper']
            if not helper.check_user_permission(screen_code, function_code, permission_type):
                helper.raise_permission_error(screen_code, function_code)

            # Execute original method
            return func(self, *args, **kwargs)

        return wrapper

    return decorator