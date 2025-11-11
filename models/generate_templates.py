# -*- coding: utf-8 -*-
"""
Excel Template Generator for Project Product Costing Module
Run this script to generate the three import templates
"""

import xlsxwriter
import os


def create_components_template():
    """Create Components Import Template"""
    workbook = xlsxwriter.Workbook('Components_Import_Template.xlsx')
    worksheet = workbook.add_worksheet('Components')

    # Formats
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#4CAF50',
        'font_color': 'white',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'font_size': 11,
    })

    example_format = workbook.add_format({
        'bg_color': '#E8F5E9',
        'border': 1,
        'align': 'left',
        'valign': 'vcenter',
    })

    number_example = workbook.add_format({
        'bg_color': '#E8F5E9',
        'border': 1,
        'align': 'right',
        'num_format': '0.00',
    })

    instruction_format = workbook.add_format({
        'bg_color': '#FFF9C4',
        'border': 1,
        'text_wrap': True,
        'valign': 'top',
        'font_size': 9,
    })

    # Set column widths
    worksheet.set_column('A:A', 35)
    worksheet.set_column('B:B', 12)
    worksheet.set_column('C:C', 15)
    worksheet.set_column('D:D', 15)
    worksheet.set_column('E:E', 20)

    # Headers
    headers = ['Component Name*', 'Quantity*', 'Weight (kg)', 'Cost Price*', 'BOM Code']
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, header_format)

    # Instructions row
    instructions = [
        'Product name or internal reference\n(Must exist in Odoo)',
        'Numeric quantity\nrequired',
        'Weight in kg\n(numeric)',
        'Unit cost price\n(numeric)',
        'Code to link with BOM\n(optional, text)'
    ]
    for col, instruction in enumerate(instructions):
        worksheet.write(1, col, instruction, instruction_format)

    # Set instruction row height
    worksheet.set_row(1, 45)

    # Example data
    examples = [
        ['Steel Sheet 2mm', 2, 5.5, 50.00, 'BOM-001'],
        ['Aluminum Plate 3mm', 4, 3.2, 75.00, 'BOM-002'],
        ['Plastic Housing ABS', 1, 0.8, 25.00, 'BOM-003'],
        ['Screws M6 Stainless', 10, 0.05, 0.50, ''],
        ['Motor Assembly 12V', 1, 2.0, 150.00, 'BOM-004'],
        ['Electronic Control Board', 1, 0.3, 85.00, 'BOM-005'],
        ['Rubber Gasket', 2, 0.1, 3.50, ''],
        ['Paint Powder Coating', 0.5, 0.5, 12.00, ''],
    ]

    for row_idx, example in enumerate(examples, start=2):
        for col_idx, value in enumerate(example):
            if col_idx in [1, 2, 3] and isinstance(value, (int, float)):
                worksheet.write(row_idx, col_idx, value, number_example)
            else:
                worksheet.write(row_idx, col_idx, value, example_format)

    # Add notes section
    notes_row = len(examples) + 4
    note_format = workbook.add_format({
        'bold': True,
        'font_color': '#D32F2F',
        'font_size': 10,
    })

    worksheet.write(notes_row, 0, '📌 IMPORTANT NOTES:', note_format)
    notes_row += 1

    notes = [
        '• Fields marked with * are required',
        '• Component Name must exactly match products in Odoo (or use internal reference)',
        '• Use decimal point (.) not comma (,) for numbers',
        '• BOM Code is optional - leave empty if no BOM needed',
        '• Delete example rows before importing your data',
        '• Keep the header row (row 1) unchanged',
    ]

    note_text_format = workbook.add_format({
        'font_size': 9,
        'text_wrap': True,
    })

    for note in notes:
        worksheet.write(notes_row, 0, note, note_text_format)
        notes_row += 1

    workbook.close()
    print("✅ Created: Components_Import_Template.xlsx")


def create_bom_materials_template():
    """Create BOM Materials Import Template"""
    workbook = xlsxwriter.Workbook('BOM_Materials_Import_Template.xlsx')
    worksheet = workbook.add_worksheet('BOM Materials')

    # Formats
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#2196F3',
        'font_color': 'white',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'font_size': 11,
    })

    example_format = workbook.add_format({
        'bg_color': '#E3F2FD',
        'border': 1,
        'align': 'left',
        'valign': 'vcenter',
    })

    number_example = workbook.add_format({
        'bg_color': '#E3F2FD',
        'border': 1,
        'align': 'right',
        'num_format': '0.00',
    })

    instruction_format = workbook.add_format({
        'bg_color': '#FFF9C4',
        'border': 1,
        'text_wrap': True,
        'valign': 'top',
        'font_size': 9,
    })

    # Set column widths
    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 35)
    worksheet.set_column('C:C', 15)
    worksheet.set_column('D:D', 15)

    # Headers
    headers = ['BOM Code*', 'Material Name*', 'Quantity*', 'Unit']
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, header_format)

    # Instructions row
    instructions = [
        'Must match BOM Code\nfrom Components',
        'Raw material product name\n(Must exist in Odoo)',
        'Quantity needed\n(numeric)',
        'Unit of measure\n(kg, pcs, meters, etc.)'
    ]
    for col, instruction in enumerate(instructions):
        worksheet.write(1, col, instruction, instruction_format)

    worksheet.set_row(1, 45)

    # Example data - organized by BOM Code
    examples = [
        ['BOM-001', 'Steel Raw Material Grade A', 6, 'kg'],
        ['BOM-001', 'Coating Material Epoxy', 0.5, 'kg'],
        ['BOM-001', 'Welding Wire', 0.2, 'kg'],
        ['', '', '', ''],
        ['BOM-002', 'Aluminum Sheet 6061', 5, 'kg'],
        ['BOM-002', 'Anodizing Chemical', 0.3, 'liter'],
        ['', '', '', ''],
        ['BOM-003', 'Plastic Pellets ABS', 1, 'kg'],
        ['BOM-003', 'Paint Black RAL9005', 0.1, 'liter'],
        ['BOM-003', 'Colorant Additive', 0.05, 'kg'],
        ['', '', '', ''],
        ['BOM-004', 'Electric Motor 12V DC', 1, 'pcs'],
        ['BOM-004', 'Wiring Harness', 1, 'set'],
        ['BOM-004', 'Mounting Bracket Steel', 2, 'pcs'],
        ['BOM-004', 'Thermal Paste', 5, 'grams'],
        ['', '', '', ''],
        ['BOM-005', 'PCB Board FR4', 1, 'pcs'],
        ['BOM-005', 'Microcontroller ATmega', 1, 'pcs'],
        ['BOM-005', 'Resistor 10K Ohm', 15, 'pcs'],
        ['BOM-005', 'Capacitor 100uF', 8, 'pcs'],
        ['BOM-005', 'LED Indicator', 3, 'pcs'],
    ]

    for row_idx, example in enumerate(examples, start=2):
        for col_idx, value in enumerate(example):
            if col_idx == 2 and isinstance(value, (int, float)) and value != '':
                worksheet.write(row_idx, col_idx, value, number_example)
            else:
                worksheet.write(row_idx, col_idx, value, example_format)

    # Add notes section
    notes_row = len(examples) + 4
    note_format = workbook.add_format({
        'bold': True,
        'font_color': '#1565C0',
        'font_size': 10,
    })

    worksheet.write(notes_row, 0, '📌 IMPORTANT NOTES:', note_format)
    notes_row += 1

    notes = [
        '• Fields marked with * are required',
        '• BOM Code must match the BOM Code from Components import',
        '• Material Name must exactly match products in Odoo',
        '• Multiple materials can have the same BOM Code',
        '• Unit is optional (e.g., kg, pcs, meters, liters)',
        '• Use decimal point (.) not comma (,) for quantities',
        '• Empty rows are ignored - use them to separate BOMs',
        '• This creates or updates BOMs with raw materials',
    ]

    note_text_format = workbook.add_format({
        'font_size': 9,
        'text_wrap': True,
    })

    for note in notes:
        worksheet.write(notes_row, 0, note, note_text_format)
        notes_row += 1

    workbook.close()
    print("✅ Created: BOM_Materials_Import_Template.xlsx")


def create_bom_operations_template():
    """Create BOM Operations Import Template"""
    workbook = xlsxwriter.Workbook('BOM_Operations_Import_Template.xlsx')
    worksheet = workbook.add_worksheet('BOM Operations')

    # Formats
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#FF9800',
        'font_color': 'white',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
        'font_size': 11,
    })

    example_format = workbook.add_format({
        'bg_color': '#FFF3E0',
        'border': 1,
        'align': 'left',
        'valign': 'vcenter',
    })

    number_example = workbook.add_format({
        'bg_color': '#FFF3E0',
        'border': 1,
        'align': 'right',
        'num_format': '0',
    })

    instruction_format = workbook.add_format({
        'bg_color': '#FFF9C4',
        'border': 1,
        'text_wrap': True,
        'valign': 'top',
        'font_size': 9,
    })

    # Set column widths
    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 30)
    worksheet.set_column('C:C', 25)
    worksheet.set_column('D:D', 20)

    # Headers
    headers = ['BOM Code*', 'Operation Name*', 'Workcenter', 'Duration (minutes)*']
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, header_format)

    # Instructions row
    instructions = [
        'Must match BOM Code\nfrom Components',
        'Name of manufacturing\noperation',
        'Workcenter name\n(Must exist in Odoo)',
        'Time in minutes\n(numeric)'
    ]
    for col, instruction in enumerate(instructions):
        worksheet.write(1, col, instruction, instruction_format)

    worksheet.set_row(1, 45)

    # Example data - organized by BOM Code with typical manufacturing operations
    examples = [
        ['BOM-001', 'Material Preparation', 'Material Storage', 5],
        ['BOM-001', 'Cutting', 'CNC Machine 1', 15],
        ['BOM-001', 'Bending', 'Press Machine', 10],
        ['BOM-001', 'Welding', 'Welding Station A', 20],
        ['BOM-001', 'Coating Application', 'Coating Line', 30],
        ['BOM-001', 'Drying', 'Drying Oven', 60],
        ['BOM-001', 'Quality Inspection', 'QC Station', 10],
        ['', '', '', ''],
        ['BOM-002', 'Material Cutting', 'CNC Machine 2', 12],
        ['BOM-002', 'Drilling', 'Drill Press', 8],
        ['BOM-002', 'Deburring', 'Finishing Station', 15],
        ['BOM-002', 'Anodizing', 'Anodizing Tank', 45],
        ['BOM-002', 'Quality Check', 'QC Station', 10],
        ['', '', '', ''],
        ['BOM-003', 'Material Loading', 'Material Storage', 3],
        ['BOM-003', 'Injection Molding', 'Molding Machine 1', 5],
        ['BOM-003', 'Cooling', 'Cooling Station', 10],
        ['BOM-003', 'Trimming', 'Trimming Station', 8],
        ['BOM-003', 'Painting', 'Paint Booth', 10],
        ['BOM-003', 'Drying', 'Drying Chamber', 30],
        ['BOM-003', 'Quality Check', 'QC Station', 5],
        ['', '', '', ''],
        ['BOM-004', 'Pre-Assembly', 'Assembly Line A', 10],
        ['BOM-004', 'Motor Installation', 'Assembly Line A', 15],
        ['BOM-004', 'Wiring', 'Assembly Line A', 20],
        ['BOM-004', 'Testing', 'Test Station', 15],
        ['BOM-004', 'Packaging', 'Packaging Area', 5],
        ['', '', '', ''],
        ['BOM-005', 'PCB Assembly', 'SMT Line', 25],
        ['BOM-005', 'Soldering', 'Soldering Station', 20],
        ['BOM-005', 'Programming', 'Programming Station', 10],
        ['BOM-005', 'Testing', 'Test Station', 15],
        ['BOM-005', 'Final Inspection', 'QC Station', 10],
    ]

    for row_idx, example in enumerate(examples, start=2):
        for col_idx, value in enumerate(example):
            if col_idx == 3 and isinstance(value, (int, float)) and value != '':
                worksheet.write(row_idx, col_idx, value, number_example)
            else:
                worksheet.write(row_idx, col_idx, value, example_format)

    # Add notes section
    notes_row = len(examples) + 4
    note_format = workbook.add_format({
        'bold': True,
        'font_color': '#E65100',
        'font_size': 10,
    })

    worksheet.write(notes_row, 0, '📌 IMPORTANT NOTES:', note_format)
    notes_row += 1

    notes = [
        '• Fields marked with * are required',
        '• BOM Code must match the BOM Code from Components import',
        '• Operation Name describes the manufacturing step',
        '• Workcenter must exist in Odoo Manufacturing module',
        '• Create workcenters in Odoo before import if they don\'t exist',
        '• Duration is in minutes (will be used for scheduling)',
        '• Operations are executed in the order listed',
        '• Empty rows are ignored - use them to separate BOMs',
        '• Multiple operations can have the same BOM Code',
    ]

    note_text_format = workbook.add_format({
        'font_size': 9,
        'text_wrap': True,
    })

    for note in notes:
        worksheet.write(notes_row, 0, note, note_text_format)
        notes_row += 1

    workbook.close()
    print("✅ Created: BOM_Operations_Import_Template.xlsx")


def create_complete_template():
    """Create Complete Import Template with all sheets"""
    workbook = xlsxwriter.Workbook('Complete_Import_Template.xlsx')

    # Sheet 1: Components
    worksheet = workbook.add_worksheet('Components')
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#4CAF50',
        'font_color': 'white',
        'border': 1,
        'align': 'center',
        'font_size': 11,
    })
    example_format = workbook.add_format({'bg_color': '#E8F5E9', 'border': 1})

    worksheet.set_column('A:A', 35)
    worksheet.set_column('B:E', 15)

    headers = ['Component Name*', 'Quantity*', 'Weight (kg)', 'Cost Price*', 'BOM Code']
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, header_format)

    examples = [
        ['Steel Sheet 2mm', 2, 5.5, 50.00, 'BOM-001'],
        ['Plastic Housing', 1, 0.8, 25.00, 'BOM-002'],
    ]
    for row_idx, example in enumerate(examples, start=1):
        for col_idx, value in enumerate(example):
            worksheet.write(row_idx, col_idx, value, example_format)

    # Sheet 2: BOM Materials
    worksheet = workbook.add_worksheet('BOM Materials')
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#2196F3',
        'font_color': 'white',
        'border': 1,
        'align': 'center',
        'font_size': 11,
    })
    example_format = workbook.add_format({'bg_color': '#E3F2FD', 'border': 1})

    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 35)
    worksheet.set_column('C:D', 15)

    headers = ['BOM Code*', 'Material Name*', 'Quantity*', 'Unit']
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, header_format)

    examples = [
        ['BOM-001', 'Steel Raw Material', 6, 'kg'],
        ['BOM-001', 'Coating Material', 0.5, 'kg'],
        ['BOM-002', 'Plastic Pellets', 1, 'kg'],
    ]
    for row_idx, example in enumerate(examples, start=1):
        for col_idx, value in enumerate(example):
            worksheet.write(row_idx, col_idx, value, example_format)

    # Sheet 3: BOM Operations
    worksheet = workbook.add_worksheet('BOM Operations')
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#FF9800',
        'font_color': 'white',
        'border': 1,
        'align': 'center',
        'font_size': 11,
    })
    example_format = workbook.add_format({'bg_color': '#FFF3E0', 'border': 1})

    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:C', 25)
    worksheet.set_column('D:D', 20)

    headers = ['BOM Code*', 'Operation Name*', 'Workcenter', 'Duration (minutes)*']
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, header_format)

    examples = [
        ['BOM-001', 'Cutting', 'CNC Machine', 15],
        ['BOM-001', 'Coating', 'Coating Line', 30],
        ['BOM-002', 'Injection Molding', 'Molding Machine', 5],
    ]
    for row_idx, example in enumerate(examples, start=1):
        for col_idx, value in enumerate(example):
            worksheet.write(row_idx, col_idx, value, example_format)

    workbook.close()
    print("✅ Created: Complete_Import_Template.xlsx")


if __name__ == '__main__':
    print("\n" + "=" * 60)
    print("📊 Excel Template Generator")
    print("   Project Product Costing Module")
    print("=" * 60 + "\n")

    print("Generating templates...\n")

    try:
        create_components_template()
        create_bom_materials_template()
        create_bom_operations_template()
        create_complete_template()

        print("\n" + "=" * 60)
        print("✅ All templates created successfully!")
        print("=" * 60)
        print("\nGenerated files:")
        print("  1. Components_Import_Template.xlsx")
        print("  2. BOM_Materials_Import_Template.xlsx")
        print("  3. BOM_Operations_Import_Template.xlsx")
        print("  4. Complete_Import_Template.xlsx")
        print("\nThese files are ready to use for importing data into Odoo.")
        print("=" * 60 + "\n")

    except Exception as e:
        print(f"\n❌ Error: {e}")
        print("Make sure xlsxwriter is installed: pip install xlsxwriter\n")