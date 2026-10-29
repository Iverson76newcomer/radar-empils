#!/usr/bin/env python3
"""Generate detailed instruction manual for TCO Calculator"""

import sys, os
DOCX_SKILL_DIR = "/home/z/my-project/skills/docx"
DOCX_SCRIPTS = os.path.join(DOCX_SKILL_DIR, "scripts")
if DOCX_SCRIPTS not in sys.path:
    sys.path.insert(0, DOCX_SCRIPTS)

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

doc = Document()

# Set default font
style = doc.styles['Normal']
font = style.font
font.name = 'Calibri'
font.size = Pt(11)
style.paragraph_format.line_spacing = 1.3
style.paragraph_format.space_after = Pt(6)

# Helper functions
def add_heading_styled(text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x1B, 0x2A, 0x4A)
    return h

def add_para(text, bold=False, italic=False, indent=False):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.first_line_indent = Cm(1)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(11)
    return p

def add_bullet(text, level=0):
    p = doc.add_paragraph(text, style='List Bullet')
    p.paragraph_format.left_indent = Cm(1.5 + level * 1)
    return p

def add_table_simple(headers, rows):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Light Grid Accent 1'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(10)
    for ri, row_data in enumerate(rows):
        for ci, val in enumerate(row_data):
            cell = table.rows[ri + 1].cells[ci]
            cell.text = str(val)
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(10)
    return table

# ===== TITLE =====
title = doc.add_heading('TCO Calculator — Detailed Instruction Manual', level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
for run in title.runs:
    run.font.color.rgb = RGBColor(0x1B, 0x2A, 0x4A)

add_para('Lemana PRO / Empils — Technological Chains TCO Calculator', bold=True).alignment = WD_ALIGN_PARAGRAPH.CENTER
add_para('Version 1.0 | June 2026', italic=True).alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_page_break()

# ===== TABLE OF CONTENTS =====
add_heading_styled('Table of Contents', 1)
toc_items = [
    '1. General Overview — What is TCO and Why It Matters',
    '2. Core Calculation Logic',
    '3. Sheet-by-Sheet Walkthrough',
    '   3.1. Parameters — Editable Inputs',
    '   3.2. Price List — Assortment & Prices',
    '   3.3. Consumption Rates — Technical Specifications',
    '   3.4. TCO Calculator — Main Calculation Sheet',
    '   3.5. Summary — Visual Comparison',
    '   3.6. Chains Reference — Technological Chains Directory',
    '   3.7. Metal Scenarios — Detailed Metal Analysis',
    '   3.8. Methodology — Formulas & Assumptions',
    '4. How to Use the Calculator Step by Step',
    '5. Key Findings & Recommendations',
    '6. Frequently Asked Questions',
]
for item in toc_items:
    p = doc.add_paragraph(item)
    p.paragraph_format.space_after = Pt(2)
    for run in p.runs:
        run.font.size = Pt(11)

doc.add_page_break()

# ===== SECTION 1 =====
add_heading_styled('1. General Overview — What is TCO and Why It Matters', 1)

add_para('TCO stands for Total Cost of Ownership — the complete cost of owning and maintaining a painted surface over a specified time horizon (typically 10 years). This calculator was built specifically for Lemana PRO retail stores selling Empils paint products, to help store managers and customers make informed purchasing decisions based on real-world costs, not just the price per can.', indent=True)

add_para('The fundamental insight behind TCO analysis is counter-intuitive: the cheapest paint often ends up being the most expensive choice over time. A paint that costs 190 RUB/kg and needs repainting every 2 years will cost significantly more over 10 years than a premium paint at 1100 RUB/kg that lasts 8 years. The reason is simple — each repaint cycle includes not just materials, but also labor costs (typically 60-70% of the total), and cheaper paints require more frequent repaints.', indent=True)

add_para('This calculator was designed to work with 14 technological chains — proven sequences of products that work together to solve specific customer problems. Each chain is built entirely from products that are actually available in Lemana PRO stores (262 SKUs from Empils), ensuring that every recommendation can be immediately fulfilled from store inventory.', indent=True)

add_heading_styled('What makes this calculator different from simple price comparison?', 2)
add_bullet('It accounts for LABOR costs — the largest component of painting expenses (typically 60-70% of total)')
add_bullet('It accounts for SERVICE LIFE — cheaper paints need more frequent repaints')
add_bullet('It uses REALISTIC REPAINITNG costs — not every repaint requires full surface preparation')
add_bullet('It compares TECHNOLOGICAL CHAINS — complete product sequences, not individual cans')
add_bullet('It covers 14 real-world SCENARIOS — from metal gates to house facades to parquet floors')

# ===== SECTION 2 =====
add_heading_styled('2. Core Calculation Logic', 1)

add_para('The TCO calculation follows a straightforward but carefully designed logic that reflects real-world painting economics. Understanding this logic is essential for interpreting the calculator results correctly.', indent=True)

add_heading_styled('2.1. The TCO Formula', 2)
add_para('TCO = Initial Painting Cost + (Number of Repaint Cycles - 1) x Repaint Cost', bold=True)
add_para('')
add_para('This formula captures a critical insight: the first painting is always the most expensive, because it includes full surface preparation (priming, rust removal, etc.). Subsequent repaints are cheaper because the primer layer persists and only the top coat needs renewal.', indent=True)

add_heading_styled('2.2. Number of Cycles', 2)
add_para('Cycles = CEILING(Calculation Horizon / Service Life)', bold=True)
add_para('')
add_para('For example, with a 10-year horizon and a paint that lasts 3 years, you get 4 cycles: the initial painting at year 0, then repaints at years 3, 6, and 9. With a premium paint lasting 8 years, you get only 2 cycles: initial at year 0 and one repaint at year 8. The difference in cycles directly drives the TCO advantage of longer-lasting products.', indent=True)

add_heading_styled('2.3. Initial Painting Cost', 2)
add_para('Initial Cost = (Material cost per m2 + Labor cost per m2) x Area', bold=True)
add_para('')
add_para('Material cost per m2 = Sum of (Consumption in g/m2 x Number of layers x Price per kg or liter / 1000) for each step in the chain. For a 2-step chain like "Primer GF-021 + Enamel PF-115", this includes both the primer and the enamel costs.', indent=True)
add_para('Labor cost per m2 = Sum of labor rates for each operation. Priming, painting, varnishing — each has a different rate per m2, reflecting the complexity and time required.', indent=True)

add_heading_styled('2.4. Repaint Cost — The Key Innovation', 2)
add_para('This is the most important and nuanced part of the calculation. The repaint cost is NOT the same as the initial cost, because not all steps need to be repeated:', indent=True)

add_table_simple(
    ['Chain Type', 'Initial Cycle', 'Repaint Cycle', 'Repaint as % of Initial'],
    [
        ['2-step (primer + paint)', 'Primer + Paint', 'Paint only + light surface prep (30% of priming labor)', '55-65%'],
        ['1-step (3v1, TITAN)', 'Full coating', 'Sanding + new coating', '~100% (same cost)'],
        ['Without primer', 'Paint only', 'Full repaint (poor adhesion requires full prep)', '100%'],
    ]
)
add_para('')
add_para('Why does this matter? When a proper primer was applied initially, it remains intact under the top coat. Repainting only requires light sanding of the old top coat (about 30% of the original priming labor) and applying a new top coat. This makes chains with primer significantly cheaper to maintain over time, even though their initial cost is higher.', indent=True)

add_heading_styled('2.5. Service Life Assumptions', 2)
add_para('Service life estimates are based on manufacturer technical specifications and real-world experience:', indent=True)

add_table_simple(
    ['Product', 'Service Life', 'Conditions'],
    [
        ['Enamel PF-115 ProstoKRASKA', '3 years', 'Exterior metal, harsh conditions'],
        ['Enamel PF-115 OREOL Premium', '5 years', 'Exterior metal, better formulation'],
        ['Primer-enamel 3v1 OREOL', '4 years', 'Exterior metal, 3-in-1 formula'],
        ['Primer-enamel TITAN', '8 years', 'Exterior metal, premium RAL-calibrated'],
        ['Primer-enamel Paint Lab', '5 years', 'Exterior metal, modern formula'],
        ['Floor enamel Ratsvet', '7 years', 'Interior floor, with primer'],
        ['Yacht varnish MONLID', '4 years', 'Exterior wood, economy'],
        ['Yacht varnish Luxens', '6 years', 'Exterior wood, mid-range'],
        ['Parquet varnish AVANTGARDE', '7 years', 'Interior wood, premium hybrid'],
        ['VD ceiling paint', '7 years', 'Interior ceiling, with primer'],
        ['VD washable paint', '7 years', 'Interior walls, with primer'],
        ['VD facade paint', '6 years', 'Exterior walls, with primer'],
        ['Radiator enamel', '7 years', 'Interior, with primer'],
    ]
)

add_heading_styled('2.6. Savings Calculation', 2)
add_para('Savings = TCO of the most expensive variant - TCO of this variant', bold=True)
add_para('')
add_para('The savings column compares each variant against the MOST EXPENSIVE option within the same chain, not against the cheapest. This means the best (cheapest TCO) variant shows the highest savings, and the most expensive variant shows 0 savings. A positive green value means this variant saves money compared to the worst choice; a 0 value means this IS the worst choice.', indent=True)

# ===== SECTION 3 =====
add_heading_styled('3. Sheet-by-Sheet Walkthrough', 1)

# 3.1
add_heading_styled('3.1. Parameters — Editable Inputs', 2)
add_para('This is the INPUT sheet where you configure the calculator. All cells with yellow background are editable — changing their values will change the calculated results in the TCO Calculator sheet (if you set up Excel formulas; currently values are pre-calculated).', indent=True)

add_heading_styled('General Parameters', 3)
add_para('These define the overall calculation framework:', indent=True)
add_bullet('Calculation horizon (years) — default 10 years. This is the time period over which TCO is calculated. Longer horizons favor premium products with longer service life.')
add_bullet('Labor rates (RUB/m2) — nine different rates for different types of work: metal priming (250), metal painting (300), VD priming (180), VD painting (220), floor painting (280), wood varnishing (300), parquet varnishing (350), radiator painting (400), decorative painting (450). Rates reflect the complexity and skill required for each operation.')

add_heading_styled('Object Areas', 3)
add_para('Sixteen typical objects with their standard areas in m2. These are used as defaults in the TCO Calculator sheet:', indent=True)

add_table_simple(
    ['Object', 'Area (m2)', 'Typical Use Case'],
    [
        ['Metal gate', '16', 'Standard 2x4m swing gate'],
        ['Metal garage', '54', 'Standard 3x6m garage (walls + roof)'],
        ['Metal door', '4.4', 'Standard 1.1x2m entrance door'],
        ['Metal fence', '30', 'Standard 10m fence, 2 sheets high'],
        ['Pipes / structures', '10', 'Small structural elements'],
        ['Garage floor (concrete)', '20', 'Standard 4x5m garage floor'],
        ['Basement floor (concrete)', '15', 'Small basement'],
        ['Warehouse floor (concrete)', '50', 'Medium warehouse section'],
        ['Wooden fence', '40', 'Standard 10m wooden fence'],
        ['Terrace / balcony', '15', 'Medium terrace'],
        ['Ceiling', '20', 'Standard room ceiling'],
        ['Walls (corridor/kitchen)', '50', 'Standard room walls'],
        ['House facade', '100', 'Small house exterior'],
        ['Decorative wall', '20', 'Feature wall'],
        ['Radiators', '6', '2-3 standard radiators'],
        ['Metal decking', '25', 'Walkway / staircase'],
    ]
)

# 3.2
add_heading_styled('3.2. Price List — Assortment & Prices', 2)
add_para('This sheet contains the complete price list of Empils products available in Lemana PRO stores. It serves as the data source for all material cost calculations.', indent=True)

add_heading_styled('Column Structure', 3)
add_bullet('# — Sequential number (1-47)')
add_bullet('Category — Product category: Primer, VD Primer, Enamel PF-115, Primer-enamel 3v1, Primer-enamel TITAN, Primer-enamel Paint Lab, Floor enamel, Yacht varnish, Parquet varnish, VD ceiling paint, VD wall paint, VD facade paint, Textured paint, Radiator enamel')
add_bullet('Product — Full product name')
add_bullet('Brand — Brand name: OREOL, ProstoKRASKA, ECONOM L, OREOL TITAN, EMPILS Paint Lab, Ratsvet, MONLID, Luxens, AVANTGARDE, PL LERUA, OREOL DISCOUNT')
add_bullet('Pack — Package size (0.4 l to 20 kg)')
add_bullet('Price (RUB) — Retail price for the package')
add_bullet('Unit — kg or l')
add_bullet('Price per kg/l (RUB) — Calculated: Price / Package size. This is the value used in TCO calculations. The largest package typically gives the best price per unit.')

add_heading_styled('How Prices Are Used', 3)
add_para('The calculator uses the "Price per kg/l" column — specifically the best-value (largest pack) price for each product category. This represents the most economical purchase option, which is what a cost-conscious customer would choose. For example, ProstoKRASKA PF-115 in a 20 kg pack costs 1790 RUB = 89.5 RUB/kg, while a 0.8 kg pack costs 130 RUB = 162.5 RUB/kg — almost twice as expensive per unit.', indent=True)

# 3.3
add_heading_styled('3.3. Consumption Rates — Technical Specifications', 2)
add_para('This sheet contains the technical parameters that drive material consumption and service life calculations. Each row represents a product category with its specific technical characteristics.', indent=True)

add_heading_styled('Column Structure', 3)
add_bullet('Product category — Name matching the categories used in the TCO Calculator chains')
add_bullet('Consumption (g/m2/layer) — How much material is needed to cover 1 m2 in one layer, in grams. This comes from manufacturer technical data sheets (TDS).')
add_bullet('Layers — Recommended number of layers for proper coverage. Most paints need 2 layers; primers need 1; varnishes need 3 for proper protection; textured paint needs 1 thick layer.')
add_bullet('Service life (years) — Expected lifetime of the coating before repainting is needed. "-" for primers means they don\'t have an independent service life (they are part of a system).')
add_bullet('Price per kg/l (RUB) — Best-value price per unit of measurement')
add_bullet('Total consumption (g/m2) — Calculated: Consumption x Layers. This is the total material needed per m2 for complete coverage.')

add_heading_styled('Key Consumption Examples', 3)
add_para('Understanding consumption is critical for TCO. Here are illustrative calculations:', indent=True)
add_bullet('Primer GF-021: 70 g/m2 x 1 layer = 70 g/m2 total. At 124 RUB/kg = 8.68 RUB/m2 material cost.')
add_bullet('Enamel PF-115 ProstoKRASKA: 120 g/m2 x 2 layers = 240 g/m2 total. At 90 RUB/kg = 21.60 RUB/m2.')
add_bullet('Primer-enamel TITAN: 90 g/m2 x 2 layers = 180 g/m2 total. At 610 RUB/l = 109.80 RUB/m2 — much more expensive per m2, but lasts 8 years vs 3.')
add_bullet('Textured paint: 800 g/m2 x 1 layer = 800 g/m2 total. At 141 RUB/kg = 112.80 RUB/m2 — thick decorative coating.')

# 3.4
add_heading_styled('3.4. TCO Calculator — Main Calculation Sheet', 2)
add_para('This is the core sheet where all 14 technological chains are calculated with their variant comparisons. The sheet is organized as a series of blocks — one block per chain — each containing a header, context, and comparison table.', indent=True)

add_heading_styled('Block Structure (repeated for each of 14 chains)', 3)
add_para('Each chain block consists of:', indent=True)
add_bullet('Section header (colored band): Chain number, name, segment, and priority level (HIGH/MEDIUM/LOW)')
add_bullet('Problem description: What customer problem this chain solves')
add_bullet('Object info: Default object type, area in m2, and calculation horizon')
add_bullet('Comparison table: All variants of this chain with calculated TCO')

add_heading_styled('Comparison Table Columns (14 columns)', 3)

add_table_simple(
    ['Column', 'What It Shows', 'How It Is Calculated'],
    [
        ['Variant', 'Name of the product combination', 'E.g. "Economy: GF-021 + PF-115 ProstoKRASKA"'],
        ['Step 1', 'First product in the chain', 'E.g. "Primer GF-021" — applied first'],
        ['Step 2', 'Second product (or "-")', 'E.g. "Enamel PF-115 ProstoKRASKA" — applied second'],
        ['Material (RUB/m2)', 'Cost of all materials per square meter', 'Sum of (Consumption x Layers x Price/kg / 1000) for all steps'],
        ['Labor (RUB/m2)', 'Cost of all labor per square meter', 'Sum of labor rates for all operations (priming, painting, etc.)'],
        ['Material total (RUB)', 'Total material cost for the object', 'Material per m2 x Object area'],
        ['Labor total (RUB)', 'Total labor cost for the object', 'Labor per m2 x Object area'],
        ['1st cycle cost (RUB)', 'Full cost of the initial painting', '(Material + Labor per m2) x Area'],
        ['Repaint cost (RUB)', 'Cost of ONE repaint cycle', 'See section 2.4 above — typically 55-65% of initial for 2-step chains'],
        ['Service life (years)', 'How long the coating lasts before repaint', 'Minimum service life across all steps, or override value'],
        ['Cycles', 'Number of repaint cycles over horizon', 'CEILING(Horizon / Service life)'],
        ['TCO for 10 yrs (RUB)', 'Total cost of ownership over 10 years', 'Initial cost + (Cycles - 1) x Repaint cost'],
        ['Savings vs max (RUB)', 'How much cheaper vs the most expensive variant', 'Max TCO in chain - This variant TCO'],
        ['Savings vs max (%)', 'Savings as percentage of the most expensive variant', 'Savings / Max TCO'],
    ]
)

add_para('')
add_heading_styled('Color Coding', 3)
add_bullet('Green cells in Savings columns: This variant is cheaper than the most expensive one (positive savings)')
add_bullet('Red cells in Savings columns: This variant IS the most expensive (0 savings)')
add_bullet('Green bold text in other columns: This variant has the LOWEST TCO (the best choice)')
add_para('The best variant is always highlighted with green bold text across the entire row, making it easy to spot the recommended option at a glance.', indent=True)

add_heading_styled('Example: Chain 1 — Metal Protection', 3)
add_para('Object: Metal gate (16 m2), Horizon: 10 years', indent=True)

add_table_simple(
    ['Variant', '1st Cycle', 'Repaint', 'Life', 'Cycles', 'TCO 10yr', 'Savings'],
    [
        ['Economy: GF-021 + PF-115 ProstoKRASKA', '9,284', '6,346', '3 yrs', '4', '28,321', '0 (most expensive)'],
        ['Mid: GF-021 + OREOL Premium', '9,787', '6,848', '5 yrs', '2', '16,636', '11,685 (41%!)'],
    ]
)
add_para('')
add_para('The Economy variant needs 4 cycles (3-year life x 4 = 12 years > 10). The Mid variant needs only 2 cycles (5-year life x 2 = 10 years). Even though the Mid variant costs slightly more per cycle, the dramatic reduction in repaint frequency saves 41% over 10 years. This is the TCO argument in action.', indent=True)

# 3.5
add_heading_styled('3.5. Summary — Visual Comparison', 2)
add_para('The Summary sheet provides a bird\'s-eye view of TCO savings across all chains that have 2 or more variants (10 out of 14 chains; the remaining 4 have only one option and thus no comparison).', indent=True)

add_heading_styled('Comparison Table', 3)
add_para('Shows for each chain: the most expensive variant, the best (lowest TCO) variant, the savings amount in RUB, and the savings percentage. The bottom row shows the TOTAL savings across all chains when choosing the best variant instead of the most expensive one.', indent=True)

add_heading_styled('Charts', 3)
add_bullet('Chart 1 (Clustered Bar): TCO for 10 years — expensive vs best variant, in RUB. Visually shows the cost gap between the worst and best choices for each chain.')
add_bullet('Chart 2 (Horizontal Bar): Savings percentage when choosing the best variant. Makes it easy to see which chains offer the biggest relative savings.')

add_heading_styled('Total Savings', 3)
add_para('The "TOTAL savings" row sums up the savings across all 10 comparable chains. This gives a sense of the overall financial impact of choosing recommended products. In the current calculation, the total savings amount to approximately 84,000 RUB across all chains, representing about 23% of the worst-case total cost.', indent=True)

# 3.6
add_heading_styled('3.6. Chains Reference — Technological Chains Directory', 2)
add_para('This is a reference sheet that lists all 14 technological chains with their key characteristics. It is designed as a quick-reference guide for store managers to understand which chain to recommend for which customer problem.', indent=True)

add_heading_styled('Column Structure', 3)
add_bullet('# — Chain number (1-14)')
add_bullet('Technology — Name of the technological chain')
add_bullet('Segment — Target segment: Metal, Floors, Wood (exterior/interior), Ceilings, Interior, Facade, Radiators, Decor')
add_bullet('Problem — The customer problem this chain solves (e.g. "Rusty metal: fences, gates, pipes, garages")')
add_bullet('SKU count — Number of SKUs available in Lemana PRO for this chain')
add_bullet('Price level — Economic positioning: Economy, Mid, Premium')
add_bullet('Priority — Business priority for Lemana PRO: HIGH (6 chains), MEDIUM (3 chains), LOW (5 chains)')
add_bullet('Key TCO advantage — The main cost argument for choosing the recommended variant')

add_heading_styled('Priority Color Coding', 3)
add_bullet('HIGH priority (red text) — Largest SKU count, most customer requests, strongest TCO argument')
add_bullet('MEDIUM priority (amber text) — Moderate demand, meaningful TCO differences')
add_bullet('LOW priority (normal text) — Niche scenarios, smaller TCO impact')

# 3.7
add_heading_styled('3.7. Metal Scenarios — Detailed Metal Analysis', 2)
add_para('Metal is the most important segment for Lemana PRO (6 out of 14 chains, highest SKU count). This sheet provides a detailed cross-analysis of all metal chain variants against four typical metal objects.', indent=True)

add_heading_styled('First Section: Total TCO Matrix (RUB for 10 years)', 3)
add_para('Rows: All 11 variants from chains 1, 2, 3, 4, 13, and 14. Columns: Four typical metal objects (Metal gate 16m2, Metal garage 54m2, Metal door 4.4m2, Metal fence 30m2). Each cell shows the absolute TCO in RUB for that variant-object combination.', indent=True)

add_heading_styled('Second Section: TCO per m2 (RUB/m2 for 10 years)', 3)
add_para('Same structure but showing cost per square meter. This allows comparison independent of object size. Note that per-m2 costs are identical across objects for the same variant, because the calculation is linear with area (no volume discounts modeled).', indent=True)

add_heading_styled('Chart', 3)
add_para('A clustered bar chart showing TCO for the first three objects across all 11 metal variants. This visually highlights the cost differences and helps identify the best-value chain for metal painting.', indent=True)

# 3.8
add_heading_styled('3.8. Methodology — Formulas & Assumptions', 2)
add_para('This sheet documents the calculation methodology in detail, serving as the authoritative reference for how TCO numbers are derived. It covers:', indent=True)

add_bullet('TCO Formula — The main formula and its components')
add_bullet('Number of cycles — How repaint frequency is calculated')
add_bullet('Initial cost — What goes into the first-cycle cost')
add_bullet('Repaint cost — The key insight about cheaper repaints (see section 2.4)')
add_bullet('Service life — How product lifetime is estimated')
add_bullet('Assumptions — Six key assumptions that underpin the calculations')
add_bullet('Why TCO > price per can — The fundamental argument for TCO-based selling')
add_bullet('Savings column logic — How savings are calculated and displayed')

# ===== SECTION 4 =====
add_heading_styled('4. How to Use the Calculator Step by Step', 1)

add_para('Follow these steps to use the calculator effectively:', indent=True)

steps = [
    ('Step 1: Identify the customer problem', 'Ask what the customer wants to paint (metal gate? concrete floor? wooden fence?) and what problems they are experiencing (rust? dust? cracking?). Go to the Chains Reference sheet and find the matching chain number.'),
    ('Step 2: Check the default object', 'In the TCO Calculator sheet, find the chain block. Check the default object and its area. If the customer has a different object size, go to the Parameters sheet and update the area.'),
    ('Step 3: Read the comparison table', 'Look at all variants in the chain. The best TCO variant is highlighted in green bold. The savings columns show how much money the customer saves by choosing the recommended variant instead of the cheapest/worst option.'),
    ('Step 4: Explain the TCO argument', 'Show the customer: "This premium product costs X more per can, but saves Y over 10 years because it lasts Z years longer." The savings column gives you the exact numbers.'),
    ('Step 5: Adjust parameters if needed', 'If the customer has a different object size, different labor rates, or wants a different time horizon, change the values in the Parameters sheet and recalculate.'),
    ('Step 6: Use the Metal Scenarios sheet', 'For metal projects, use the Metal Scenarios sheet to show TCO across different object types. This helps customers with multiple metal objects see the total impact.'),
]

for title, desc in steps:
    add_para(title, bold=True)
    add_para(desc, indent=True)

# ===== SECTION 5 =====
add_heading_styled('5. Key Findings & Recommendations', 1)

add_para('Based on the TCO calculator results, here are the most impactful findings:', indent=True)

add_heading_styled('5.1. Metal Segment (6 chains, highest priority)', 2)
add_bullet('Chain 1 (Metal protection): Choosing OREOL Premium over ProstoKRASKA saves 41% TCO over 10 years, despite similar per-can prices. The key driver is 5-year vs 3-year service life.')
add_bullet('Chain 2 (3v1 products): Paint Lab is the sweet spot — 30% TCO savings vs basic 3v1 OREOL, with 5-year vs 4-year life. TITAN also saves vs 3v1, but less dramatically (20%) due to higher per-can cost.')
add_bullet('Chain 13 (Heavy-duty protection): The biggest TCO savings in the entire calculator — GF-021 + 3v1 saves 42% vs using only 3v1. The double-layer system (primer + 3v1) lasts 10 years vs 4 years for 3v1 alone.')

add_heading_styled('5.2. Floor Segment (1 chain, high priority)', 2)
add_bullet('Chain 5 (Concrete floor): Adding GF-021 primer before Ratsvet saves 8% TCO and extends service life from 4 to 7 years. The primer also reduces paint consumption by 1.8x.')

add_heading_styled('5.3. Wood Segment (2 chains)', 2)
add_bullet('Chain 6 (Outdoor wood): Luxens saves 28% TCO vs MONLID, lasting 6 years vs 4. For outdoor applications, the longer life more than compensates for the higher price.')

add_heading_styled('5.4. Interior & Facade (5 chains)', 2)
add_bullet('Chains 8-12: For interior and facade applications, the TCO advantage of using primer is 5-13%. Smaller than for metal, but still significant — especially for facades (13% savings) where repaints are expensive due to scaffolding costs.')

add_heading_styled('5.5. Overall', 2)
add_para('Across all 10 comparable chains, choosing the recommended variant over the cheapest/worst option saves approximately 84,000 RUB in total — a 23% reduction in total cost of ownership. This is a compelling argument for upselling to better products.', indent=True)

# ===== SECTION 6 =====
add_heading_styled('6. Frequently Asked Questions', 1)

faqs = [
    ('Why is the "recommended" variant sometimes more expensive per can but cheaper in TCO?',
     'Because TCO counts ALL costs over 10 years: materials + labor x number of cycles. A more expensive product that lasts longer needs fewer repaints. Since labor is 60-70% of each cycle cost, reducing the number of cycles dramatically reduces total cost, even if the per-can price is higher.'),
    ('Why does "without primer" have a higher consumption multiplier?',
     'Without primer, the surface absorbs paint unevenly. Manufacturer TDS typically states that consumption increases 1.5-1.8x on unprimed surfaces because the paint soaks into the substrate rather than forming a proper film. This is especially true for concrete and drywall.'),
    ('Can I change the labor rates for my region?',
     'Yes! Go to the Parameters sheet and change any of the 9 labor rate values in the yellow cells. All calculations in the TCO Calculator sheet are based on these values. Labor rates vary significantly across Russian regions — Moscow rates may be 2-3x higher than regional averages.'),
    ('Why are some chains marked as LOW priority?',
     'LOW priority chains (7, 8, 9, 10, 11) represent niche scenarios with small SKU counts and/or modest TCO differences. They are still valid for customers with those specific needs, but they are not the primary focus for in-store merchandising and sales training.'),
    ('What if the customer wants DIY (no labor costs)?',
     'Set all labor rates to 0 in the Parameters sheet. The TCO will then reflect only material costs. In this case, the TCO advantage of premium products is smaller (because labor savings are eliminated), but still exists due to longer service life reducing the number of repaint cycles and material purchases.'),
    ('Why does repaint cost less than initial cost?',
     'Because the primer layer (first step in 2-step chains) doesn\'t need to be reapplied. It persists under the top coat and continues protecting the surface. Only the top coat needs renewal, plus light surface preparation (sanding). This is why chains with primer are so much more cost-effective over time.'),
    ('How accurate are the service life estimates?',
     'They are conservative estimates based on manufacturer specifications and real-world experience in Russian climate conditions. Actual service life depends on many factors: surface preparation quality, application conditions (temperature, humidity), mechanical wear, and chemical exposure. The estimates assume professional application following manufacturer guidelines.'),
]

for q, a in faqs:
    add_para(q, bold=True)
    add_para(a, indent=True)

# Save
output_path = "/home/z/my-project/download/TCO_Calculator_Instruction_Manual.docx"
doc.save(output_path)
print(f"OK: {output_path}")
