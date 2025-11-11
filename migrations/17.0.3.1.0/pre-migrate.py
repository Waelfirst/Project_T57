# -*- coding: utf-8 -*-

def migrate(cr, version):
    """Migration script to create specification tables"""
    
    # Create component_specification_definition table
    cr.execute("""
        CREATE TABLE IF NOT EXISTS component_specification_definition (
            id SERIAL PRIMARY KEY,
            name VARCHAR NOT NULL,
            code VARCHAR,
            sequence INTEGER DEFAULT 10,
            active BOOLEAN DEFAULT TRUE,
            description TEXT,
            create_uid INTEGER,
            create_date TIMESTAMP,
            write_uid INTEGER,
            write_date TIMESTAMP
        )
    """)
    
    # Create component_specification_value table
    cr.execute("""
        CREATE TABLE IF NOT EXISTS component_specification_value (
            id SERIAL PRIMARY KEY,
            sequence INTEGER DEFAULT 10,
            pricing_component_id INTEGER REFERENCES project_product_component(id) ON DELETE CASCADE,
            planning_component_id INTEGER REFERENCES material_planning_component(id) ON DELETE CASCADE,
            specification_id INTEGER REFERENCES component_specification_definition(id) ON DELETE RESTRICT,
            specification_name VARCHAR,
            value VARCHAR NOT NULL,
            notes TEXT,
            create_uid INTEGER,
            create_date TIMESTAMP,
            write_uid INTEGER,
            write_date TIMESTAMP
        )
    """)
    
    # Create indexes
    cr.execute("""
        CREATE INDEX IF NOT EXISTS idx_spec_value_pricing 
        ON component_specification_value(pricing_component_id)
    """)
    
    cr.execute("""
        CREATE INDEX IF NOT EXISTS idx_spec_value_planning 
        ON component_specification_value(planning_component_id)
    """)
    
    cr.execute("""
        CREATE INDEX IF NOT EXISTS idx_spec_value_spec 
        ON component_specification_value(specification_id)
    """)
