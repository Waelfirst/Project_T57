# -*- coding: utf-8 -*-

import logging

_logger = logging.getLogger(__name__)


def post_init_hook(env):
    """Create default specification definitions after module installation"""
    
    _logger.info('Creating default specification definitions...')
    
    SpecDef = env['component.specification.definition']
    
    # Check if specifications already exist
    existing = SpecDef.search([('code', 'in', [
        'MATERIAL', 'DIMENSIONS', 'COLOR', 'FINISH', 'THICKNESS',
        'WEIGHT_CAP', 'QUALITY', 'CERT', 'SUPPLIER', 'NOTES'
    ])])
    
    if existing:
        _logger.info('Default specifications already exist, skipping creation')
        return
    
    # Create default specifications
    specifications = [
        {'name': 'Material', 'code': 'MATERIAL', 'sequence': 10},
        {'name': 'Dimensions', 'code': 'DIMENSIONS', 'sequence': 20},
        {'name': 'Color', 'code': 'COLOR', 'sequence': 30},
        {'name': 'Finish', 'code': 'FINISH', 'sequence': 40},
        {'name': 'Thickness', 'code': 'THICKNESS', 'sequence': 50},
        {'name': 'Weight Capacity', 'code': 'WEIGHT_CAP', 'sequence': 60},
        {'name': 'Quality Grade', 'code': 'QUALITY', 'sequence': 70},
        {'name': 'Certification', 'code': 'CERT', 'sequence': 80},
        {'name': 'Preferred Supplier', 'code': 'SUPPLIER', 'sequence': 90},
        {'name': 'Special Notes', 'code': 'NOTES', 'sequence': 100},
    ]
    
    for spec_data in specifications:
        try:
            SpecDef.create(spec_data)
            _logger.info('Created specification: %s', spec_data['name'])
        except Exception as e:
            _logger.error('Error creating specification %s: %s', spec_data['name'], e)
    
    _logger.info('Default specifications created successfully')
