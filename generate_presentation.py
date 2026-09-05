from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Create presentation object
prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(7.5)

# Define color scheme
TITLE_COLOR = RGBColor(25, 75, 150)
ACCENT_COLOR = RGBColor(220, 20, 60)
TEXT_COLOR = RGBColor(50, 50, 50)

def add_title_slide(prs, title, subtitle):
    """Add a title slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(230, 240, 250)
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2.5), Inches(9), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    p = title_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(54)
    p.font.bold = True
    p.font.color.rgb = TITLE_COLOR
    p.alignment = PP_ALIGN.CENTER
    
    # Add subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.2), Inches(9), Inches(1))
    subtitle_frame = subtitle_box.text_frame
    p = subtitle_frame.paragraphs[0]
    p.text = subtitle
    p.font.size = Pt(28)
    p.font.color.rgb = ACCENT_COLOR
    p.alignment = PP_ALIGN.CENTER

def add_content_slide(prs, title, content_list):
    """Add a content slide with bullet points"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = TITLE_COLOR
    
    # Add horizontal line under title
    line = slide.shapes.add_connector(1, Inches(0.5), Inches(1.15), Inches(9.5), Inches(1.15))
    line.line.color.rgb = ACCENT_COLOR
    line.line.width = Pt(2)
    
    # Add content
    content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(8.4), Inches(5.5))
    text_frame = content_box.text_frame
    text_frame.word_wrap = True
    
    for i, item in enumerate(content_list):
        if i == 0:
            p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()
        p.text = item
        p.font.size = Pt(18)
        p.font.color.rgb = TEXT_COLOR
        p.level = 0
        p.space_before = Pt(6)
        p.space_after = Pt(6)

# Slide 1: Title Slide
add_title_slide(prs, "Buckingham Pi Theorem", "A Comprehensive Study Guide")

# Slide 2: Introduction
add_content_slide(prs, "What is the Buckingham Pi Theorem?", [
    "• A fundamental principle in dimensional analysis",
    "• Named after physicist Edgar Buckingham (1914)",
    "• Provides a systematic method to reduce complex physical problems",
    "• Converts dimensional quantities into dimensionless parameters",
    "• Essential tool in engineering and physics research"
])

# Slide 3: Basic Concept
add_content_slide(prs, "Basic Concept", [
    "• States that physically meaningful equations must be dimensionally consistent",
    "• If an equation relates n physical quantities with m independent dimensions,",
    "   the equation can be reduced to a relationship of (n - m) dimensionless groups",
    "• These dimensionless groups are called 'Pi groups' or 'Buckingham Pi numbers'",
    "• Formula: Number of Pi terms = n - m"
])

# Slide 4: Mathematical Foundation
add_content_slide(prs, "Mathematical Foundation", [
    "• Let: n = total number of variables",
    "         m = number of independent dimensions (fundamental dimensions)",
    "         k = number of dimensionless groups (Pi terms)",
    "• Then: k = n - m",
    "• Each Pi term is dimensionless: [Π] = [M⁰ L⁰ T⁰ ...]",
    "• The relationship can be expressed as: Π₁ = f(Π₂, Π₃, ..., Πₖ)"
])

# Slide 5: Fundamental Dimensions
add_content_slide(prs, "Fundamental Dimensions", [
    "• Mass (M)",
    "• Length (L)",
    "• Time (T)",
    "• Temperature (Θ)",
    "• Electric Current (A)",
    "• Amount of Substance (N)",
    "• Luminous Intensity (J)"
])

# Slide 6: Steps to Apply Buckingham Pi Theorem
add_content_slide(prs, "Steps to Apply the Theorem", [
    "Step 1: Identify all relevant variables and their dimensions",
    "Step 2: Determine the number of independent dimensions (m)",
    "Step 3: Calculate k = n - m (number of Pi groups needed)",
    "Step 4: Select m repeating variables (dimensionally independent)",
    "Step 5: Form Pi groups by combining repeating variables with other variables",
    "Step 6: Verify that each Pi group is dimensionless",
    "Step 7: Express the functional relationship between Pi groups"
])

# Slide 7: Selecting Repeating Variables
add_content_slide(prs, "Selecting Repeating Variables", [
    "• Must be dimensionally independent",
    "• Should not form a dimensionless group themselves",
    "• Should be common to most variables in the problem",
    "• Typically include:",
    "   - A dimension of length or area",
    "   - A dimension of mass or force",
    "   - A dimension of time or velocity",
    "• Number of repeating variables = m (number of fundamental dimensions)"
])

# Slide 8: Example Problem Setup
add_content_slide(prs, "Example: Drag Force on a Sphere", [
    "Problem: Find how drag force F depends on:",
    "   • Velocity of fluid (v) [L T⁻¹]",
    "   • Diameter of sphere (d) [L]",
    "   • Density of fluid (ρ) [M L⁻³]",
    "   • Viscosity of fluid (μ) [M L⁻¹ T⁻¹]",
    "Variables: n = 5 (F, v, d, ρ, μ)",
    "Dimensions: m = 3 (M, L, T)",
    "Pi groups needed: k = 5 - 3 = 2"
])

# Slide 9: Example Solution - Forming Pi Groups
add_content_slide(prs, "Example Solution - Pi Groups", [
    "Π₁ = F / (ρ v² d²)  [Drag coefficient]",
    "Π₂ = (μ) / (ρ v d)  [Reynolds number]",
    "General relationship:",
    "   Π₁ = f(Π₂)",
    "   F / (ρ v² d²) = f(ρ v d / μ)",
    "   F = ρ v² d² × f(Re)"
])

# Slide 10: Common Dimensionless Numbers
add_content_slide(prs, "Common Dimensionless Numbers", [
    "• Reynolds Number (Re) = ρ v L / μ  [Flow characteristics]",
    "• Froude Number (Fr) = v / √(g L)  [Wave/gravity effects]",
    "• Mach Number (Ma) = v / a  [Compressibility effects]",
    "• Strouhal Number (Sr) = f L / v  [Oscillatory flows]",
    "• Nusselt Number (Nu) = h L / k  [Heat transfer]",
    "• Prandtl Number (Pr) = c μ / k  [Thermal properties]"
])

# Slide 11: Applications
add_content_slide(prs, "Applications of Buckingham Pi Theorem", [
    "• Fluid Mechanics: Understanding flow patterns and drag",
    "• Heat Transfer: Design of heat exchangers",
    "• Structural Analysis: Scaling laws for structures",
    "• Aerodynamics: Aircraft design and wind tunnel testing",
    "• Chemical Engineering: Reactor design and scaling",
    "• Material Science: Testing material properties at different scales"
])

# Slide 12: Advantages
add_content_slide(prs, "Advantages", [
    "✓ Reduces the number of variables to study",
    "✓ Guides experimental design efficiently",
    "✓ Enables scaling from models to prototypes",
    "✓ Provides physical insight into problems",
    "✓ Saves time and resources in research",
    "✓ Applicable across diverse fields"
])

# Slide 13: Limitations
add_content_slide(prs, "Limitations", [
    "• Does not provide the exact form of the relationship",
    "• Requires correct identification of all relevant variables",
    "• Missing variables can lead to incomplete analysis",
    "• Doesn't reveal constants or coefficients",
    "• Experimental or theoretical work still needed to determine f(Π)",
    "• Not suitable for problems with non-physical variables"
])

# Slide 14: Tips for Problem Solving
add_content_slide(prs, "Tips for Problem Solving", [
    "1. List ALL relevant variables (don't miss any!)",
    "2. Determine dimensions carefully and correctly",
    "3. Choose repeating variables that appear in multiple terms",
    "4. Verify dimensionless nature of each Pi group",
    "5. Cross-check your work by substituting back",
    "6. Interpret physical meaning of dimensionless groups",
    "7. Use standard dimensionless numbers when applicable"
])

# Slide 15: Summary
add_content_slide(prs, "Summary", [
    "• Buckingham Pi Theorem: k = n - m",
    "• Systematically reduces dimensional analysis problems",
    "• Creates dimensionless groups (Pi terms)",
    "• Essential for experimental design and scaling",
    "• Widely used across engineering and science disciplines",
    "• Provides foundation for similarity and modeling theory"
])

# Slide 16: References & Further Reading
add_content_slide(prs, "References & Further Reading", [
    "• Buckingham, E. (1914) 'On Physically Similar Systems'",
    "• White, F.M. (2008) 'Fluid Mechanics' - McGraw-Hill",
    "• Barenblatt, G.I. (1996) 'Scaling, Self-similarity, and Intermediate Asymptotics'",
    "• Sonin, A.A. 'The Physical Basis of Dimensional Analysis' - MIT",
    "• Szirtes, T. & Rózsa, P. (2006) 'Applied Dimensional Analysis'",
    "• Visit: https://en.wikipedia.org/wiki/Buckingham_π_theorem"
])

# Save presentation
prs.save('Buckingham_Pi_Theorem_Study_Notes.pptx')
print("✓ Presentation created successfully!")
print("✓ File saved as: Buckingham_Pi_Theorem_Study_Notes.pptx")
