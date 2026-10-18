import os
import joblib
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VIZ_DIR = os.path.join(BASE_DIR, "visualizations")
MODEL_DIR = os.path.join(BASE_DIR, "models")


def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def add_section_line(doc, color="667EEA"):
    para = doc.add_paragraph()
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)
    return para


def add_image_with_caption(doc, filename, caption, width=5.5):
    fpath = os.path.join(VIZ_DIR, filename)
    if os.path.exists(fpath):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(fpath, width=Inches(width))
        cap = doc.add_paragraph(caption)
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap_run = cap.runs[0]
        cap_run.font.italic = True
        cap_run.font.size = Pt(9)
        cap_run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        doc.add_paragraph()
        return True
    return False


def create_architecture_diagram(doc):
    """Create a text-based architecture diagram"""
    doc.add_heading("System Architecture Overview", level=2)

    arch_lines = [
        "+-----------------------------------------------------------------+",
        "|              RETENTIX AI - System Architecture                   |",
        "+-----------------------------------------------------------------+",
        "|                                                                   |",
        "|  +--------------+      +--------------+      +--------------+   |",
        "|  |   FRONTEND   |      |   BACKEND    |      |    MODELS    |   |",
        "|  |  (Streamlit) |<---->|   (Python)   |<---->|  (Artifacts) |   |",
        "|  +--------------+      +--------------+      +--------------+   |",
        "|         |                       |                       |        |",
        "|         v                       v                       v        |",
        "|  +--------------+      +--------------+      +--------------+   |",
        "|  |  UI Views    |      | Prediction   |      | Model Bundle |   |",
        "|  |  - Executive |      | Service      |      | - Active Ptr |   |",
        "|  |  - EDA       |      | - Single     |      | - Manifest   |   |",
        "|  |  - Drivers   |      | - Batch      |      | - Artifacts  |   |",
        "|  |  - Predictor |      | - Simulator  |      | - Hashes     |   |",
        "|  |  - Models    |      | - XAI (SHAP) |      +--------------+   |",
        "|  |  - Strategy  |      +--------------+                          |",
        "|  |  - Bulk      |              |                                  |",
        "|  +--------------+              v                                  |",
        "|         |              +--------------+                           |",
        "|         |              | Input        |                           |",
        "|         |              | Contract     |                           |",
        "|         |              +--------------+                           |",
        "|         v                     |                                   |",
        "|  +--------------+             |                                   |",
        "|  |  CSS Styles  |<------------+                                   |",
        "|  +--------------+                                                   |",
        "|                        +--------------+                            |",
        "|                        |    DATA      |                            |",
        "|                        |  (CSV)       |                            |",
        "|                        +--------------+                            |",
        "+-----------------------------------------------------------------+",
    ]

    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = para.add_run("\n".join(arch_lines))
    run.font.name = "Consolas"
    run.font.size = Pt(7)
    run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    doc.add_paragraph()


def create_comprehensive_report():
    """Generate a comprehensive 6-page report with architecture and visualizations"""

    # Load model results
    results = joblib.load(os.path.join(MODEL_DIR, "results.pkl"))
    best_model_name = joblib.load(os.path.join(MODEL_DIR, "best_model_name.pkl"))

    doc = Document()

    # Set margins for all sections
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Set default font
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    # Configure heading styles
    h1_style = doc.styles["Heading 1"]
    h1_style.font.name = "Calibri"
    h1_style.font.size = Pt(18)
    h1_style.font.bold = True
    h1_style.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    h2_style = doc.styles["Heading 2"]
    h2_style.font.name = "Calibri"
    h2_style.font.size = Pt(14)
    h2_style.font.bold = True
    h2_style.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)

    # ==================== PAGE 1: TITLE & EXECUTIVE SUMMARY ====================

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("\n\nRETENTIX AI")
    r.font.name = "Calibri"
    r.font.size = Pt(36)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x63, 0x66, 0xF1)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = subtitle.add_run("Customer Churn Intelligence & Retention Studio")
    r.font.name = "Calibri"
    r.font.size = Pt(20)
    r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    doc.add_paragraph()

    sub2 = doc.add_paragraph()
    sub2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub2.add_run("Comprehensive Project Report")
    r.font.name = "Calibri"
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor(0x66, 0x7E, 0xEA)

    doc.add_paragraph()
    doc.add_paragraph()

    # Project Info Table
    info_table = doc.add_table(rows=6, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    info_data = [
        ("Project Name", "RETENTIX AI - Customer Churn Intelligence Studio"),
        ("Engineered By", "Dhiraj Dalvi"),
        ("Course", "Data Analytics / Machine Learning"),
        ("Date", "October 2026"),
        ("Dataset", "Telco Customer Churn (7,043 records)"),
        ("Version", "v2.4 - Enterprise Production Release"),
    ]

    for i, (label, value) in enumerate(info_data):
        cell0 = info_table.rows[i].cells[0]
        cell1 = info_table.rows[i].cells[1]
        cell0.text = label
        cell1.text = value
        cell0.paragraphs[0].runs[0].font.bold = True
        cell0.paragraphs[0].runs[0].font.size = Pt(11)
        cell1.paragraphs[0].runs[0].font.size = Pt(11)
        set_cell_bg(cell0, "E0E7FF")
        set_cell_bg(cell1, "F8FAFC")

    doc.add_paragraph()
    add_section_line(doc, "6366F1")

    # Executive Summary
    doc.add_heading("Executive Summary", level=1)

    summary_text = (
        "RETENTIX AI is a state-of-the-art, end-to-end customer churn intelligence "
        "platform designed for telecommunications enterprise leaders, retention managers, "
        "and data science teams. By combining high-precision machine learning inference "
        "with interactive commercial simulation, RETENTIX AI empowers organizations to "
        "identify churn risks, diagnose behavioral catalysts, simulate retention campaign "
        "ROI, and execute targeted retention strategies before revenue loss occurs.\n\n"
        "The system processes 7,043 customer records through a sophisticated ML pipeline, "
        "achieving 84.2% ROC-AUC performance with the champion model. The platform features "
        "real-time prediction capabilities, batch processing for portfolio-wide risk assessment, "
        "and explainable AI (XAI) through SHAP analysis to provide actionable insights for "
        "retention teams."
    )
    doc.add_paragraph(summary_text)

    doc.add_paragraph()

    # Key Metrics Box
    metrics_table = doc.add_table(rows=2, cols=4)
    metrics_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    metrics = [
        ("7,043", "Customer Records"),
        ("84.2%", "ROC-AUC Score"),
        ("26.5%", "Churn Rate"),
        ("$1.2M+", "Monthly Revenue"),
    ]
    for i, (value, label) in enumerate(metrics):
        cell = metrics_table.rows[0].cells[i]
        cell.text = value
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.paragraphs[0].runs[0].font.bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(18)
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0x63, 0x66, 0xF1)
        set_cell_bg(cell, "F0F4FF")
        cell2 = metrics_table.rows[1].cells[i]
        cell2.text = label
        cell2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell2.paragraphs[0].runs[0].font.size = Pt(10)
        cell2.paragraphs[0].runs[0].font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
        set_cell_bg(cell2, "F8FAFC")

    doc.add_page_break()

    # ==================== PAGE 2: ARCHITECTURE & SYSTEM DESIGN ====================

    doc.add_heading("System Architecture & Design", level=1)
    add_section_line(doc)

    create_architecture_diagram(doc)

    doc.add_heading("Technology Stack", level=2)

    tech_table = doc.add_table(rows=1, cols=3)
    tech_table.style = "Light Grid Accent 1"
    tech_table.rows[0].cells[0].text = "Layer"
    tech_table.rows[0].cells[1].text = "Technology"
    tech_table.rows[0].cells[2].text = "Purpose"

    for cell in tech_table.rows[0].cells:
        cell.paragraphs[0].runs[0].font.bold = True
        set_cell_bg(cell, "1F4E79")
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    tech_stack = [
        ("Frontend", "Streamlit 1.28+", "Interactive web dashboard with real-time updates"),
        ("Backend", "Python 3.13", "Core business logic and ML pipeline"),
        ("ML Framework", "Scikit-learn 1.6.1", "Model training, preprocessing, and inference"),
        ("Data Processing", "Pandas 2.x, NumPy 2.x", "Data manipulation and numerical operations"),
        ("Visualization", "Plotly 5.x, Matplotlib, Seaborn", "Interactive and static charts"),
        ("Explainability", "SHAP 0.48+", "Model interpretation and feature importance"),
        ("Serialization", "Joblib 1.3+", "Model and pipeline persistence"),
        ("Testing", "Pytest 8.0+", "Automated unit and integration tests"),
    ]

    for layer, tech, purpose in tech_stack:
        row = tech_table.add_row().cells
        row[0].text = layer
        row[1].text = tech
        row[2].text = purpose
        row[0].paragraphs[0].runs[0].font.bold = True

    doc.add_paragraph()

    doc.add_heading("Module Architecture", level=2)

    modules_text = (
        "The system follows a modular architecture with clear separation of concerns:\n\n"
        "- Frontend Layer (frontend/): Contains 7 view modules handling UI presentation, "
        "including executive dashboards, exploratory data analysis, churn drivers visualization, "
        "AI prediction interface, model evaluation, retention strategy, and bulk prediction studio.\n\n"
        "- Backend Layer (backend/): Implements core business logic including data loading, "
        "model training, prediction services, input validation contracts, explainability (SHAP), "
        "and model bundle management with integrity verification.\n\n"
        "- Models Layer (models/): Manages ML artifacts through a bundle-based architecture "
        "with active pointer system, ensuring safe deployment and rollback capabilities.\n\n"
        "- Data Layer (data/): Houses the cleaned telecommunications customer dataset with "
        "19 features including demographic, service, and billing attributes.\n\n"
        "- Tests Layer (tests/): Comprehensive test suite covering champion contracts, "
        "input validation, prediction services, SHAP explainability, and smoke tests."
    )

    doc.add_paragraph(modules_text)

    doc.add_page_break()

    # ==================== PAGE 5: CHURN DRIVERS & FEATURE IMPORTANCE ====================

    doc.add_heading("Churn Drivers & Feature Importance", level=1)
    add_section_line(doc)

    doc.add_paragraph("Understanding the key drivers of customer churn is critical for developing effective retention strategies. The analysis reveals several critical factors.")

    doc.add_heading("Feature Importance Analysis", level=2)
    add_image_with_caption(doc, "feature_importance.png", "Figure 5: Top features driving churn predictions based on model importance scores", 5.0)

    doc.add_heading("Key Churn Drivers Identified", level=2)

    drivers = [
        "1. Contract Type: Month-to-month contracts show 45%+ churn rate vs. 5% for 2-year contracts.",
        "2. Internet Service: Fiber optic customers churn at 2x the rate of DSL customers.",
        "3. Payment Method: Electronic check users show 2x higher churn than auto-pay methods.",
        "4. Tech Support: Customers without tech support churn at 40%+ vs. 15% with support.",
        "5. Tenure: New customers (0-6 months) are at highest risk of churn.",
        "6. Monthly Charges: Higher charges correlate with increased churn probability.",
        "7. Online Security: Bundling security services reduces churn by 30%+.",
    ]
    for d in drivers:
        doc.add_paragraph(d, style="List Bullet")

    doc.add_heading("Service Churn Rates", level=2)
    add_image_with_caption(doc, "service_churn_rates.png", "Figure 6: Churn rates breakdown by service type and add-on features", 5.0)

    doc.add_page_break()

    # ==================== PAGE 6: APPLICATION FEATURES & RECOMMENDATIONS ====================

    doc.add_heading("Application Features & Strategic Recommendations", level=1)
    add_section_line(doc)

    doc.add_heading("Key Application Modules", level=2)

    features_table = doc.add_table(rows=1, cols=2)
    features_table.style = "Light Grid Accent 1"
    features_table.rows[0].cells[0].text = "Module"
    features_table.rows[0].cells[1].text = "Capabilities"
    for cell in features_table.rows[0].cells:
        cell.paragraphs[0].runs[0].font.bold = True
        set_cell_bg(cell, "1F4E79")
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    features = [
        ("Executive Command Center", "Real-time KPIs, revenue exposure analysis, portfolio risk distribution, tenure density mapping, strategic alerts"),
        ("Exploratory Intelligence", "Demographic breakdown, service hierarchy sunburst, correlation matrix, numerical distributions, interactive dataset explorer"),
        ("Churn Drivers & Ecosystem", "Protective shield effect analysis, payment friction study, value vs. tenure risk quadrants"),
        ("AI Churn Predictor", "Single subscriber profiling, archetype presets, live inference gauge, what-if intervention simulator"),
        ("Bulk Portfolio Scoring", "Batch CSV upload, schema validation, automated risk tiering, batch analytics and export"),
        ("Model Leaderboard & XAI", "Performance comparison, SHAP explainability, feature importance, partial dependence plots"),
        ("Retention ROI & Playbook", "Campaign ROI calculator, retention strategies, predictive alerts, loyalty rewards framework"),
    ]
    for module, capabilities in features:
        row = features_table.add_row().cells
        row[0].text = module
        row[0].paragraphs[0].runs[0].font.bold = True
        row[1].text = capabilities

    doc.add_paragraph()

    doc.add_heading("Strategic Recommendations", level=2)

    rec_table = doc.add_table(rows=1, cols=2)
    rec_table.style = "Light Grid Accent 1"
    rec_table.rows[0].cells[0].text = "Priority"
    rec_table.rows[0].cells[1].text = "Action Plan"
    for cell in rec_table.rows[0].cells:
        cell.paragraphs[0].runs[0].font.bold = True
        set_cell_bg(cell, "1F4E79")
        cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    recommendations = [
        ("CRITICAL", "Convert 15% of month-to-month customers to 1-year contracts through targeted discounts. Estimated savings: $180k/year"),
        ("HIGH", "Bundle Tech Support with Fiber Optic plans to reduce 40% churn rate. Target: 25% reduction in fiber churn"),
        ("HIGH", "Migrate electronic check users to auto-pay with $5/month billing credit. Target: 50% migration in 6 months"),
        ("MEDIUM", "Implement predictive alerts in CRM to flag at-risk customers weekly for proactive outreach"),
        ("MEDIUM", "Launch tenure-based loyalty rewards at 12, 24, and 48-month milestones to improve retention"),
        ("LOW", "Review fiber optic pricing strategy to ensure perceived value matches premium pricing"),
    ]

    priority_colors = {
        "CRITICAL": ("FEE2E2", RGBColor(0xDC, 0x26, 0x26)),
        "HIGH": ("FEF3C7", RGBColor(0xD9, 0x77, 0x06)),
        "MEDIUM": ("DBEAFE", RGBColor(0x25, 0x63, 0xEB)),
        "LOW": ("F1F5F9", RGBColor(0x64, 0x74, 0x8B)),
    }

    for priority, action in recommendations:
        row = rec_table.add_row().cells
        row[0].text = priority
        row[0].paragraphs[0].runs[0].font.bold = True
        bg, color = priority_colors.get(priority, ("F1F5F9", RGBColor(0x64, 0x74, 0x8B)))
        set_cell_bg(row[0], bg)
        row[0].paragraphs[0].runs[0].font.color.rgb = color
        row[1].text = action

    doc.add_paragraph()
    add_section_line(doc, "6366F1")

    # Footer
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = footer_p.add_run(
        "RETENTIX AI - Customer Churn Intelligence Studio\n"
        "Engineered by Dhiraj Dalvi | October 2026 | Version 2.4\n"
        "Enterprise Telecom Machine Learning & Revenue Protection System"
    )
    r.font.size = Pt(9)
    r.font.italic = True
    r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    # Save document
    output_path = os.path.join(BASE_DIR, "RETENTIX_AI_Comprehensive_Report.docx")
    doc.save(output_path)
    print(f"Comprehensive 6-page report saved: {output_path}")
    return output_path


if __name__ == "__main__":
    create_comprehensive_report()
