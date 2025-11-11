# -*- coding: utf-8 -*-
{
    'name': 'Project Product Planning & Costing Management',
    'version': '17.0.3.2.5',
    'category': 'Project',
    'summary': 'Manage projects, product costing, and material planning',
    'description': """
        Project Product Planning & Costing Management
        ==============================================
        * Define projects with finished products
        * Version-based product pricing and components
        * Material and production planning
        * BOM integration and work order generation
    """,
    'author': 'Your Company',
    'website': 'https://www.yourcompany.com',
    'depends': [
        'base',
        'project',
        'product',
        'stock',
        'mrp',
        'purchase',
        'sale_management',
    ],
    'data': [
        'security/ir.model.access.csv',
        'data/sequence_data.xml',
        'data/screen_definitions_data.xml',  # ← جديد
        'views/project_definition_views.xml',
        'views/project_product_pricing_views.xml',
        'views/material_production_planning_views.xml',
        'views/work_order_execution_views.xml',
        'views/work_order_process_wizard_views.xml',  # ← ADD THIS
        'views/production_report_views.xml',
        'views/component_specification_views.xml',
        'views/import_wizard_views.xml',
        'views/work_order_wizard_views.xml',
        'views/template_generator_wizard_views.xml',  # ← ADD THIS
        'views/user_permission_views.xml',  # ← جديد
        'views/project_cost_estimation_views.xml',  # ← إضافة
        'views/menu_views.xml',
    ],
    'demo': [],
    'installable': True,
    'application': True,
    'auto_install': False,
    'license': 'LGPL-3',
    'post_init_hook': 'post_init_hook',
}
