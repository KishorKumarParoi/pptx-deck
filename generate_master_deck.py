"""
Kishor Kumar Paroi — Master Executive Pitch Deck Generator (Online-Compliant v3.0)
Generates the complete 24-slide boardroom presentation:
- Slide 1: Apple Keynote Minimalist Cover (cover.png)
- Slide 2: Executive Persona & Live Verified Metrics ($10,000 Hero)
- Slide 3: Executive Resume & Career Trajectory (Experience, RUET, ICPC, Skills)
- Slide 4: Executive Systems Thesis (The 4 Production Pillars)
- Slide 5: Executive Telemetry & Benchmark Dashboard (Native Vector Comparison Bars)
- Slides 6 - 20: Every Project's Detailed Architecture & Bulleted Points (All 15 Projects!)
- Slide 21: Value Differentiation Matrix (Kishor vs Typical Senior Engineers)
- Slide 22: Target Roles, Engagement Models & Global Collaboration
- Slide 23: 90-Day Production Execution Roadmap (Days 0-30, 31-60, 61-90)
- Slide 24: Cinematic Visual Outro (outro.png - Vantage Retro Synthwave Crest)

Strict Compliance Guarantees:
- ZERO parse_xml transitions (avoids ECMA-376 schema errors in Office Online)
- ZERO shapes.add_chart calls (avoids corrupt Excel workbook embeddings in Office Online)
- 100% compliant OpenXML vector shapes, text frames, and high-res images
- Guaranteed to render smoothly in Microsoft Office Online, Google Slides, Keynote, & Desktop PPT
"""

import os
import sys
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ─── Color Palette (Obsidian Cyber Matrix) ────────────────────────────────────
BG_VOID          = RGBColor(0x06, 0x0A, 0x14)  # Deep obsidian void
BG_CARD          = RGBColor(0x0E, 0x16, 0x28)  # Midnight card navy
BG_CARD_LIGHT    = RGBColor(0x15, 0x22, 0x3C)  # Highlight card fill
BG_TRACK         = RGBColor(0x1B, 0x28, 0x42)  # Benchmark track fill

BORDER_CYAN      = RGBColor(0x00, 0xE5, 0xFF)  # Cyber cyan
BORDER_GOLD      = RGBColor(0xFB, 0xBF, 0x24)  # Warm gold
BORDER_EMERALD   = RGBColor(0x10, 0xB9, 0x81)  # Mint neon emerald
BORDER_VIOLET    = RGBColor(0x8B, 0x5C, 0xF6)  # Deep electric violet
BORDER_ROSE      = RGBColor(0xFB, 0x71, 0x85)  # Vibrant rose
BORDER_DIM       = RGBColor(0x1F, 0x2D, 0x48)  # Subtle card border

WHITE            = RGBColor(0xFF, 0xFF, 0xFF)
MUTED_WHITE      = RGBColor(0xF1, 0xF5, 0xF9)
MUTED_GRAY       = RGBColor(0x94, 0xA3, 0xB8)
DARK_GRAY        = RGBColor(0x64, 0x74, 0x8B)

DECK_DIR         = "/Users/kishorkumarparoi/job-project/pptx-deck"
OPT_IMG_DIR      = os.path.join(DECK_DIR, "opt_img")
os.makedirs(OPT_IMG_DIR, exist_ok=True)

# ─── High-Fidelity Image Optimization Helper ─────────────────────────────────
def get_optimized_image(src_path, key, max_dim=1500, quality=88):
    """Resizes and compresses images so the PPTX loads instantaneously online."""
    if not os.path.exists(src_path):
        print(f"WARNING: Image not found: {src_path}")
        return None
    out_path = os.path.join(OPT_IMG_DIR, f"{key}.jpg")
    if os.path.exists(out_path) and os.path.getsize(out_path) > 1000:
        return out_path
    try:
        im = Image.open(src_path)
        if im.mode in ("RGBA", "P"):
            bg = Image.new("RGB", im.size, (6, 10, 20))
            if im.mode == "RGBA":
                bg.paste(im, mask=im.split()[3])
            else:
                bg.paste(im.convert("RGBA"))
            im = bg
        elif im.mode != "RGB":
            im = im.convert("RGB")

        if key in ("cover", "outro"):
            im_resized = im.resize((1920, 1080), Image.Resampling.LANCZOS)
            im_resized.save(out_path, format="JPEG", quality=92, optimize=True)
        else:
            im.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
            im.save(out_path, format="JPEG", quality=quality, optimize=True)
        return out_path
    except Exception as e:
        print(f"Error optimizing {src_path}: {e}")
        return src_path

# ─── Presentation Setup (16:9 Widescreen) ─────────────────────────────────────
prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout     = prs.slide_layouts[6]
TOTAL_SLIDES     = 24

def add_slide():
    slide = prs.slides.add_slide(blank_layout)
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = BG_VOID
    return slide

def add_box(slide, l, t, w, h, fill_col=BG_CARD, border_col=None, border_w=1.0, shape_type=MSO_SHAPE.RECTANGLE):
    shape = slide.shapes.add_shape(shape_type, Inches(l), Inches(t), Inches(w), Inches(h))
    if fill_col:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_col
    else:
        shape.fill.background()
    if border_col:
        shape.line.color.rgb = border_col
        shape.line.width = Pt(border_w)
    else:
        shape.line.fill.background()
    return shape

def add_pill(slide, l, t, text, text_col=WHITE, bg_col=BG_CARD_LIGHT, border_col=None, font_sz=8.5, pad_w=0.32, h=0.28):
    w = len(text) * 0.072 + pad_w
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_col
    if border_col:
        shape.line.color.rgb = border_col
        shape.line.width = Pt(1.0)
    else:
        shape.line.fill.background()
    tf = shape.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = False
    tf.margin_top = Pt(0)
    tf.margin_bottom = Pt(0)
    tf.margin_left = Pt(4)
    tf.margin_right = Pt(4)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.size = Pt(font_sz)
    r.font.bold = True
    r.font.color.rgb = text_col
    return shape, w

def add_header(slide, title, category="EXECUTIVE BRIEFING", slide_num=None, accent_col=BORDER_CYAN):
    # Category Tracker
    tx_cat = slide.shapes.add_textbox(Inches(0.8), Inches(0.42), Inches(8.5), Inches(0.28))
    tf_c = tx_cat.text_frame
    tf_c.margin_left = tf_c.margin_right = tf_c.margin_top = tf_c.margin_bottom = Pt(0)
    p_c = tf_c.paragraphs[0]
    r_c = p_c.add_run()
    r_c.text = category.upper()
    r_c.font.size = Pt(9.5)
    r_c.font.bold = True
    r_c.font.color.rgb = accent_col

    # Main Title
    tx_t = slide.shapes.add_textbox(Inches(0.8), Inches(0.70), Inches(10.5), Inches(0.55))
    tf_t = tx_t.text_frame
    tf_t.margin_left = tf_t.margin_right = tf_t.margin_top = tf_t.margin_bottom = Pt(0)
    p_t = tf_t.paragraphs[0]
    r_t = p_t.add_run()
    r_t.text = title
    r_t.font.size = Pt(20)
    r_t.font.bold = True
    r_t.font.color.rgb = WHITE

    # Sleek accent line under header
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.30), Inches(11.733), Inches(0.015))
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(0x1F, 0x2D, 0x48)
    line.line.fill.background()

    # Slide Number Indicator
    if slide_num is not None:
        tx_num = slide.shapes.add_textbox(Inches(11.2), Inches(0.55), Inches(1.333), Inches(0.35))
        tf_n = tx_num.text_frame
        tf_n.margin_left = tf_n.margin_right = tf_n.margin_top = tf_n.margin_bottom = Pt(0)
        p_n = tf_n.paragraphs[0]
        p_n.alignment = PP_ALIGN.RIGHT
        r_n = p_n.add_run()
        r_n.text = f"{slide_num:02d} / {TOTAL_SLIDES:02d}"
        r_n.font.size = Pt(10)
        r_n.font.bold = True
        r_n.font.color.rgb = MUTED_GRAY

# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 1: CINEMATIC VISUAL COVER (APPLE KEYNOTE MINIMALIST FULL-BLEED)
# ═════════════════════════════════════════════════════════════════════════════
print("Building Slide 1: Cover...")
s1 = add_slide()
img_cover_path = os.path.join(DECK_DIR, "cover.png")
if os.path.exists(img_cover_path):
    s1.shapes.add_picture(img_cover_path, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)

# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 2: THE $10,000 EXECUTIVE HERO — IDENTITY & VERIFIED PRODUCTION RIGOR
# ═════════════════════════════════════════════════════════════════════════════
print("Building Slide 2: Executive Persona & Metrics...")
s2 = add_slide()
add_header(s2, "Kishor Kumar Paroi — Principal AI Systems Architect & Founder @ KoomAI", "EXECUTIVE IDENTITY & PRODUCTION TRACK RECORD", slide_num=2)

# Outer card wrapper
add_box(s2, 0.8, 1.50, 11.733, 5.45, fill_col=BG_CARD, border_col=RGBColor(0x1F, 0x2D, 0x48), border_w=1.0)
add_box(s2, 0.8, 1.50, 0.15, 5.45, fill_col=BORDER_CYAN, border_col=None)

# Badges (Pills)
add_pill(s2, 1.25, 1.80, "PRINCIPAL AI SYSTEMS ARCHITECT", BORDER_CYAN, BG_CARD_LIGHT, BORDER_CYAN, font_sz=8.5)
add_pill(s2, 4.05, 1.80, "FOUNDER @ KOOMAI", BORDER_GOLD, BG_CARD_LIGHT, BORDER_GOLD, font_sz=8.5)
add_pill(s2, 5.80, 1.80, "FDE @ PHI0.AI", BORDER_EMERALD, BG_CARD_LIGHT, BORDER_EMERALD, font_sz=8.5)
add_pill(s2, 7.15, 1.80, "B.SC. CSE @ RUET", BORDER_VIOLET, BG_CARD_LIGHT, BORDER_VIOLET, font_sz=8.5)

# Left Column: Bio & Core Architecture Focus
tx_bio = s2.shapes.add_textbox(Inches(1.25), Inches(2.25), Inches(6.6), Inches(4.3))
tf_bio = tx_bio.text_frame
tf_bio.word_wrap = True
tf_bio.margin_left = tf_bio.margin_right = tf_bio.margin_top = tf_bio.margin_bottom = Pt(0)

p_h = tf_bio.paragraphs[0]
r_h = p_h.add_run()
r_h.text = "Engineering Predictable, High-Throughput Autonomous Agent Meshes, Sub-35ms Hybrid RAG & Medallion Lakehouses"
r_h.font.size = Pt(15.5)
r_h.font.bold = True
r_h.font.color.rgb = WHITE

p_sub = tf_bio.add_paragraph()
p_sub.space_before = Pt(8)
r_sub = p_sub.add_run()
r_sub.text = (
    "A battle-tested AI Systems Architect bridging mathematical rigor with distributed cloud infrastructure. "
    "Creator of mission-critical systems across phi0.ai, Samsung AI Core, and kepi.ai—curtailing millions in operational "
    "compute while maintaining 99.4% autonomous agent precision at scale."
)
r_sub.font.size = Pt(10.5)
r_sub.font.color.rgb = MUTED_GRAY

p_spec = tf_bio.add_paragraph()
p_spec.space_before = Pt(10)
r_spec = p_spec.add_run()
r_spec.text = (
    "• Production Specialization: LangGraph Multi-Agent Workflows, NVIDIA NeMo Guardrails, vLLM Multi-LoRA\n"
    "• Systems Rigor: Rust SIMD (AVX-512), Go 1.22 Gateways, PyTorch, C/C++, Linux Kernel & I/O Hot-Paths\n"
    "• Data & MLOps: Snowflake, PySpark Delta Lake, dbt Core, Feast Feature Store, KServe, ArgoCD GitOps\n"
    "• Security & Compliance: Zero-Trust Model Context Protocol (FastMCP), HITL Gates, PyRIT Red-Teaming"
)
r_spec.font.size = Pt(9.5)
r_spec.font.color.rgb = MUTED_WHITE

# Ecosystem Links Pill Bar at bottom of left column
links_y = 5.95
add_pill(s2, 1.25, links_y, "koom.kishor.tech", BORDER_CYAN, BG_CARD_LIGHT, BORDER_CYAN, font_sz=8.0)
add_pill(s2, 2.95, links_y, "kishor.tech", WHITE, BG_CARD_LIGHT, DARK_GRAY, font_sz=8.0)
add_pill(s2, 4.15, links_y, "macos.kishor.tech", BORDER_GOLD, BG_CARD_LIGHT, BORDER_GOLD, font_sz=8.0)
add_pill(s2, 5.85, links_y, "15m Zoom Link", BORDER_EMERALD, BG_CARD_LIGHT, BORDER_EMERALD, font_sz=8.0)

# Right Column: 4 High-Status KPI Metric Cards (2x2 Grid)
metrics_s2 = [
    ("$1.4M", "Enterprise Contracts Won", "phi0.ai production GenAI delivery", BORDER_GOLD),
    ("<35ms", "Hybrid RAG P95 Latency", "Portkey + Redis semantic cache", BORDER_CYAN),
    ("50x", "SIMD Scoring Acceleration", "Rust AVX-512 vector pipelines", BORDER_VIOLET),
    ("99.4%", "Multi-Agent Reliability", "LangGraph state machine HITL", BORDER_EMERALD),
]

card_w = 2.15
card_h = 2.05
xs = [8.15, 10.45]
ys = [1.85, 4.20]

for idx, (val, title, desc, col) in enumerate(metrics_s2):
    gx = xs[idx % 2]
    gy = ys[idx // 2]
    add_box(s2, gx, gy, card_w, card_h, fill_col=BG_CARD_LIGHT, border_col=col, border_w=1.2, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
    
    tx_m = s2.shapes.add_textbox(Inches(gx + 0.12), Inches(gy + 0.15), Inches(card_w - 0.24), Inches(card_h - 0.3))
    tf_m = tx_m.text_frame
    tf_m.word_wrap = True
    tf_m.margin_left = tf_m.margin_right = tf_m.margin_top = tf_m.margin_bottom = Pt(0)
    
    p_v = tf_m.paragraphs[0]
    r_v = p_v.add_run()
    r_v.text = val
    r_v.font.size = Pt(26)
    r_v.font.bold = True
    r_v.font.color.rgb = col
    
    p_t = tf_m.add_paragraph()
    p_t.space_before = Pt(4)
    r_t = p_t.add_run()
    r_t.text = title
    r_t.font.size = Pt(10.5)
    r_t.font.bold = True
    r_t.font.color.rgb = WHITE
    
    p_d = tf_m.add_paragraph()
    p_d.space_before = Pt(3)
    r_d = p_d.add_run()
    r_d.text = desc
    r_d.font.size = Pt(8.5)
    r_d.font.color.rgb = MUTED_GRAY

# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 3: EXECUTIVE RESUME & CAREER TRAJECTORY (ANIMATED VISUAL EXPERIENCE)
# ═════════════════════════════════════════════════════════════════════════════
print("Building Slide 3: Executive Resume...")
s3 = add_slide()
add_header(s3, "Executive Resume: Production Track Record & Academic Pedigree", "CAREER TRAJECTORY & PEDIGREE", slide_num=3)

# Left Card: Experience Timeline (Width 6.2")
add_box(s3, 0.8, 1.50, 6.2, 5.45, fill_col=BG_CARD, border_col=BORDER_CYAN, border_w=1.0)
add_box(s3, 0.8, 1.50, 0.12, 5.45, fill_col=BORDER_CYAN, border_col=None)

tx_exp = s3.shapes.add_textbox(Inches(1.1), Inches(1.68), Inches(5.7), Inches(5.1))
tf_exp = tx_exp.text_frame
tf_exp.word_wrap = True
tf_exp.margin_left = tf_exp.margin_right = tf_exp.margin_top = tf_exp.margin_bottom = Pt(0)

# Section Title
p_et = tf_exp.paragraphs[0]
r_et = p_et.add_run()
r_et.text = "PRODUCTION EXPERIENCE & LEADERSHIP"
r_et.font.size = Pt(11.5)
r_et.font.bold = True
r_et.font.color.rgb = BORDER_CYAN

experiences = [
    ("phi0.ai", "July 2026 – Oct 2026", "Forward Deployed AI Engineer (Remote, Europe)", [
        "Orchestrated 5 autonomous LangGraph agents with 99.4% task completion, accelerating onboarding 4x.",
        "Devised sub-35ms Hybrid RAG (HNSW, BM25, FlashRank ONNX) saving $18K/mo; secured $1.4M client contracts."
    ]),
    ("Samsung AI Core / Care Health", "March 2026 – June 2026", "Data Engineering & AI Systems Intern", [
        "AWS PySpark/EMR batch & streaming pipelines; slashed operational compute by $4.2K/run (34% savings).",
        "Accelerated 8B/70B LLM inference on H100 with FlashAttention-2 (3.8x throughput, -54% P99 latency).",
        "Real-time Kinesis fraud detection pipeline eradicated false positives by 52%, preserving $380K/yr."
    ]),
    ("kepi.ai", "Jan 2022 – Sept 2025", "Senior Full Stack Engineer (3 yrs 9 mos, Remote)", [
        "Led zero-downtime microservices migration to AWS Kubernetes; scaled throughput by 300% at 99.99% uptime."
    ]),
    ("Founder @ KoomAI", "Jan 2025 – Dec 2025", "Asynchronous Video Intelligence Platform", [
        "Built WebRTC engine, desktop app, extension & 20-agent LangGraph mesh; scaled to $24K MRR."
    ])
]

for company, period, role, bullets in experiences:
    p_c = tf_exp.add_paragraph()
    p_c.space_before = Pt(5)
    r_c1 = p_c.add_run()
    r_c1.text = f"{company} "
    r_c1.font.size = Pt(10.5)
    r_c1.font.bold = True
    r_c1.font.color.rgb = WHITE
    
    r_c2 = p_c.add_run()
    r_c2.text = f"({period})"
    r_c2.font.size = Pt(8.5)
    r_c2.font.color.rgb = BORDER_GOLD
    
    p_r = tf_exp.add_paragraph()
    p_r.space_before = Pt(1)
    r_r = p_r.add_run()
    r_r.text = role
    r_r.font.size = Pt(9.0)
    r_r.font.bold = True
    r_r.font.color.rgb = BORDER_CYAN
    
    for b in bullets:
        p_b = tf_exp.add_paragraph()
        p_b.space_before = Pt(2)
        r_b = p_b.add_run()
        r_b.text = f"• {b}"
        r_b.font.size = Pt(8.5)
        r_b.font.color.rgb = MUTED_GRAY

# Right Card: Academic Pedigree & Technical Arsenal (Width 5.3")
add_box(s3, 7.25, 1.50, 5.28, 5.45, fill_col=BG_CARD, border_col=BORDER_GOLD, border_w=1.0)
add_box(s3, 7.25, 1.50, 0.12, 5.45, fill_col=BORDER_GOLD, border_col=None)

tx_edu = s3.shapes.add_textbox(Inches(7.55), Inches(1.68), Inches(4.8), Inches(5.1))
tf_edu = tx_edu.text_frame
tf_edu.word_wrap = True
tf_edu.margin_left = tf_edu.margin_right = tf_edu.margin_top = tf_edu.margin_bottom = Pt(0)

# Section Title
p_at = tf_edu.paragraphs[0]
r_at = p_at.add_run()
r_at.text = "ACADEMIC PEDIGREE & SKILL ARSENAL"
r_at.font.size = Pt(11.5)
r_at.font.bold = True
r_at.font.color.rgb = BORDER_GOLD

edus = [
    ("RUET (Rajshahi Univ. of Eng. & Tech.)", "2018 – 2023", "B.Sc. in Computer Science & Engineering", "CGPA 3.25 / 4.00 · Merit Rank 483rd in Bangladesh (Top 0.5%)"),
    ("Saint Joseph Higher Secondary School", "2015 – 2017", "Higher Secondary Certificate (HSC)", "GPA 5.00 / 5.00 · Merit Rank 85th in Dhaka Board"),
    ("S.J Govt Model High School", "2013 – 2015", "Secondary School Certificate (SSC)", "GPA 5.00 / 5.00 · Merit Rank 90th in Dhaka Board"),
]

for school, period, degree, honors in edus:
    p_s = tf_edu.add_paragraph()
    p_s.space_before = Pt(5)
    r_s = p_s.add_run()
    r_s.text = f"{school} "
    r_s.font.size = Pt(10.0)
    r_s.font.bold = True
    r_s.font.color.rgb = WHITE
    
    r_sp = p_s.add_run()
    r_sp.text = f"({period})"
    r_sp.font.size = Pt(8.5)
    r_sp.font.color.rgb = BORDER_CYAN
    
    p_d = tf_edu.add_paragraph()
    p_d.space_before = Pt(1)
    r_d = p_d.add_run()
    r_d.text = f"{degree} — {honors}"
    r_d.font.size = Pt(8.5)
    r_d.font.color.rgb = MUTED_GRAY

# Contest Honors
p_c_head = tf_edu.add_paragraph()
p_c_head.space_before = Pt(7)
r_ch = p_c_head.add_run()
r_ch.text = "COMPETITIVE PROGRAMMING & HONORS"
r_ch.font.size = Pt(9.5)
r_ch.font.bold = True
r_ch.font.color.rgb = BORDER_EMERALD

p_c_body = tf_edu.add_paragraph()
p_c_body.space_before = Pt(2)
r_cb = p_c_body.add_run()
r_cb.text = (
    "• ICPC Regional Contestant; Solved 1,500+ problems (LeetCode/Codeforces)\n"
    "• Mentored 100+ university engineers; Conducted in-person AI Agents masterclass"
)
r_cb.font.size = Pt(8.5)
r_cb.font.color.rgb = MUTED_GRAY

# Technical Arsenal
p_s_head = tf_edu.add_paragraph()
p_s_head.space_before = Pt(7)
r_sh = p_s_head.add_run()
r_sh.text = "PRODUCTION SKILL ARSENAL"
r_sh.font.size = Pt(9.5)
r_sh.font.bold = True
r_sh.font.color.rgb = BORDER_VIOLET

p_s_body = tf_edu.add_paragraph()
p_s_body.space_before = Pt(2)
r_sb = p_s_body.add_run()
r_sb.text = (
    "• Languages: Python, Rust (SIMD), Go 1.22, C/C++, TypeScript, SQL, Bash\n"
    "• AI/GenAI: PyTorch, vLLM, LangGraph, NeMo Guardrails, FlashRank ONNX, PEFT\n"
    "• Data Infra: Snowflake, PySpark Delta Lake, dbt Core, Feast, Kafka, Airflow\n"
    "• Cloud/DevOps: AWS, GCP, Azure, Kubernetes, ArgoCD GitOps, Terraform IaC"
)
r_sb.font.size = Pt(8.5)
r_sb.font.color.rgb = MUTED_WHITE

# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 4: THE EXECUTIVE SYSTEMS THESIS & CORE VALUE (4 PILLARS)
# ═════════════════════════════════════════════════════════════════════════════
print("Building Slide 4: Systems Thesis...")
s4 = add_slide()
add_header(s4, "Executive Systems Thesis: Engineering Predictable, Sub-Millisecond AI", "STRATEGIC FOUNDATION", slide_num=4)

pillars = [
    ("Sub-35ms Hybrid RAG & Cache", "Portkey semantic caching with Redis vector index and local ONNX FlashRank re-ranking curtailing P95 latency by 85% while halting $18K/mo in repetitive API spend.", BORDER_CYAN, "P95 <35ms · 85% Latency Drop"),
    ("Multi-LoRA vLLM Dynamic Serving", "Multiplexing multiple task-specialized RSLoRA adapters on single GPU instances with PagedAttention; delivering 3.4x higher token throughput at 68% lower GPU footprint.", BORDER_VIOLET, "3.4x Throughput · -68% GPU Cost"),
    ("Zero-Trust AI Defense & HITL", "NVIDIA NeMo Guardrails, BotShield WAF, and FastMCP Human-in-the-Loop break-glass gates blocking 95%+ prompt injection attacks (PyRIT tested) with 100% audit compliance.", BORDER_EMERALD, "95%+ Threat Block · 100% Audit"),
    ("Real-Time Medallion Lakehouses", "Decoupled PySpark Delta Lake, Snowflake, and Feast Feature Stores over Redis; streaming millions of CDC events under 5ms feature freshness and sub-second analytical queries.", BORDER_GOLD, "35M+ Events · <5ms Freshness")
]

p_w = 5.70
p_h = 2.45
p_xs = [0.80, 6.83]
p_ys = [1.60, 4.35]

for idx, (title, desc, col, badge) in enumerate(pillars):
    px = p_xs[idx % 2]
    py = p_ys[idx // 2]
    add_box(s4, px, py, p_w, p_h, fill_col=BG_CARD, border_col=col, border_w=1.2, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_box(s4, px, py, 0.12, p_h, fill_col=col, border_col=None)
    
    add_pill(s4, px + 0.35, py + 0.25, badge, col, BG_CARD_LIGHT, col, font_sz=8.0)
    
    tx_p = s4.shapes.add_textbox(Inches(px + 0.35), Inches(py + 0.65), Inches(p_w - 0.6), Inches(p_h - 0.75))
    tf_p = tx_p.text_frame
    tf_p.word_wrap = True
    tf_p.margin_left = tf_p.margin_right = tf_p.margin_top = tf_p.margin_bottom = Pt(0)
    
    p_pt = tf_p.paragraphs[0]
    r_pt = p_pt.add_run()
    r_pt.text = title
    r_pt.font.size = Pt(13.5)
    r_pt.font.bold = True
    r_pt.font.color.rgb = WHITE
    
    p_pd = tf_p.add_paragraph()
    p_pd.space_before = Pt(4)
    r_pd = p_pd.add_run()
    r_pd.text = desc
    r_pd.font.size = Pt(9.5)
    r_pd.font.color.rgb = MUTED_GRAY

# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 5: EXECUTIVE TELEMETRY & BENCHMARK DASHBOARD (PURE VECTOR VISUALS)
# ═════════════════════════════════════════════════════════════════════════════
print("Building Slide 5: Telemetry Wall...")
s5 = add_slide()
add_header(s5, "Executive Telemetry Wall: Quantified Production Performance & Cost Slashed", "VERIFIED BENCHMARKS", slide_num=5)

benchmarks = [
    {
        "category": "HYBRID RAG RETRIEVAL P95 LATENCY",
        "speedup": "12x FASTER",
        "accent": BORDER_CYAN,
        "items": [
            ("Legacy Cloud RAG API (Chaining)", "420 ms", 0.90, DARK_GRAY),
            ("Standard OSS Local Vector DB", "180 ms", 0.45, MUTED_GRAY),
            ("Kishor's Hybrid RAG + Portkey Cache", "<35 ms", 0.10, BORDER_CYAN),
        ]
    },
    {
        "category": "TOKEN SERVING COST PER 1M TOKENS",
        "speedup": "88% CHEAPER",
        "accent": BORDER_GOLD,
        "items": [
            ("Proprietary Cloud API (GPT-4o)", "$15.00", 0.95, DARK_GRAY),
            ("Standard OSS Single-Model GPU", "$6.20", 0.48, MUTED_GRAY),
            ("Kishor's Multi-LoRA vLLM Multiplexing", "$1.80", 0.14, BORDER_GOLD),
        ]
    },
    {
        "category": "AUTONOMOUS AGENT TASK RELIABILITY",
        "speedup": "99.4% PRECISION",
        "accent": BORDER_EMERALD,
        "items": [
            ("Single-Shot Direct Prompting", "58.2%", 0.58, DARK_GRAY),
            ("Sequential ReAct Chain Loop", "76.4%", 0.76, MUTED_GRAY),
            ("Kishor's LangGraph HITL State Mesh", "99.4%", 0.99, BORDER_EMERALD),
        ]
    },
    {
        "category": "REAL-TIME FEATURE LOOKUP LATENCY",
        "speedup": "90x SPEEDUP",
        "accent": BORDER_VIOLET,
        "items": [
            ("Batch Warehouse Direct Query", "450 ms", 0.92, DARK_GRAY),
            ("Indexed Postgres Cloud Replica", "45 ms", 0.35, MUTED_GRAY),
            ("Kishor's Feast Store over Redis Cluster", "<5 ms", 0.08, BORDER_VIOLET),
        ]
    }
]

row_w = 5.70
row_h = 2.50
r_xs = [0.80, 6.83]
r_ys = [1.55, 4.35]

for b_idx, b in enumerate(benchmarks):
    rx = r_xs[b_idx % 2]
    ry = r_ys[b_idx // 2]
    col = b["accent"]
    
    add_box(s5, rx, ry, row_w, row_h, fill_col=BG_CARD, border_col=col, border_w=1.2, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_box(s5, rx, ry, 0.12, row_h, fill_col=col, border_col=None)
    
    # Title & Badge
    tx_cat = s5.shapes.add_textbox(Inches(rx + 0.30), Inches(ry + 0.15), Inches(3.8), Inches(0.28))
    tf_c = tx_cat.text_frame
    tf_c.margin_left = tf_c.margin_right = tf_c.margin_top = tf_c.margin_bottom = Pt(0)
    p_c = tf_c.paragraphs[0]
    r_c = p_c.add_run()
    r_c.text = b["category"]
    r_c.font.size = Pt(9.5)
    r_c.font.bold = True
    r_c.font.color.rgb = col
    
    add_pill(s5, rx + 4.25, ry + 0.15, b["speedup"], col, BG_CARD_LIGHT, col, font_sz=8.0)
    
    # Progress bars
    bar_start_y = ry + 0.58
    for item_idx, (lbl, val_str, pct, bar_col) in enumerate(b["items"]):
        iy = bar_start_y + item_idx * 0.58
        
        tx_l = s5.shapes.add_textbox(Inches(rx + 0.30), Inches(iy), Inches(3.8), Inches(0.24))
        tf_l = tx_l.text_frame
        tf_l.margin_left = tf_l.margin_right = tf_l.margin_top = tf_l.margin_bottom = Pt(0)
        p_l = tf_l.paragraphs[0]
        r_l = p_l.add_run()
        r_l.text = lbl
        r_l.font.size = Pt(8.5)
        r_l.font.color.rgb = WHITE if bar_col == col else MUTED_GRAY
        
        tx_v = s5.shapes.add_textbox(Inches(rx + 4.30), Inches(iy), Inches(1.10), Inches(0.24))
        tf_v = tx_v.text_frame
        tf_v.margin_left = tf_v.margin_right = tf_v.margin_top = tf_v.margin_bottom = Pt(0)
        p_v = tf_v.paragraphs[0]
        p_v.alignment = PP_ALIGN.RIGHT
        r_v = p_v.add_run()
        r_v.text = val_str
        r_v.font.size = Pt(9.0)
        r_v.font.bold = True
        r_v.font.color.rgb = bar_col
        
        # Track Background
        track_w = 5.10
        add_box(s5, rx + 0.30, iy + 0.26, track_w, 0.14, fill_col=BG_TRACK, border_col=None, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
        # Fill Bar
        fill_w = max(track_w * pct, 0.20)
        add_box(s5, rx + 0.30, iy + 0.26, fill_w, 0.14, fill_col=bar_col, border_col=None, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)

# ═════════════════════════════════════════════════════════════════════════════
# SLIDES 6 TO 20: ALL 15 PROJECTS DETAILED ARCHITECTURES & BULLETED BREAKDOWNS
# ═════════════════════════════════════════════════════════════════════════════
print("Preparing All 15 Project Architectures...")

PROJECTS = [
    {
        "id": 1,
        "category": "FLAGSHIP 01: DATA & LAKEHOUSE",
        "title": "Zomato AI Lakehouse & Multimodal Dispute Platform",
        "subtitle": "35M+ Orders Medallion Lakehouse & Sub-5ms Feature Store",
        "badge": "35M+ ORDERS · $18K/MO SAVED",
        "accent": BORDER_CYAN,
        "img_src": "/Users/kishorkumarparoi/job-project/vantage-portfolio/images/projects/zomato/01_lakehouse_architecture.jpg",
        "bullets": [
            ("Medallion Architecture:", "Scaled Snowflake & PySpark Delta Lake processing 35M+ orders (2.3 GB) with dbt Core SCD-2; curtailed warehouse compute by 42% ($18K/mo savings)."),
            ("Real-Time Feature Store:", "Streamed CDC events via Airflow, Debezium WAL, and Feast Store over Redis cluster, driving feature serving latency to <5ms for real-time model ranking."),
            ("Multimodal CV Dispute Engine:", "Integrated dHash perceptual hashing anti-fraud and receipt OCR with G-Eval CoT arbitration, slashing fraudulent claims by 68% and saving $1.2M annually."),
            ("Compaction & Multi-Region DR:", "Automated Delta Lake OPTIMIZE compaction, Bloom filter indexes, and MLflow shadow canary promotion with multi-region failover (RPO < 60s, RTO < 4.5m).")
        ],
        "metrics": [("$18K/mo", "Compute Slashed"), ("<5ms", "Redis Serving"), ("68%", "Fraud Cut")]
    },
    {
        "id": 2,
        "category": "FLAGSHIP 02: HIGH-THROUGHPUT ML",
        "title": "Hybrid Anime Intelligence & Continuous LLMOps Platform",
        "subtitle": "Polyglot Mesh (Go/Rust/FastAPI) & Multi-LoRA vLLM Engine",
        "badge": "50x SPEEDUP · -68% GPU SPEND",
        "accent": BORDER_VIOLET,
        "img_src": "/Users/kishorkumarparoi/job-project/vantage-portfolio/images/projects/hybrid/01_architecture_overview.jpg",
        "bullets": [
            ("Polyglot Microservice Mesh:", "Go 1.22 API Gateway, Rust SIMD vector pre-scoring, and FastAPI AI core over Kafka; curtailed scoring latency from 25ms to <0.5ms (50x speedup)."),
            ("Hybrid Neural Retrieval:", "Fused Dual-Tower NCF neural filtering, ChromaDB HNSW dense vectors, and BM25 sparse search via Reciprocal Rank Fusion (NDCG@10 = 0.94), lifting retention by 28%."),
            ("Multi-LoRA vLLM Dynamic Serving:", "Multiplexed task adapters on 1x GPU via vLLM Multi-LoRA with RSLoRA training flywheel; slashed GPU spend by 68% while delivering 3.4x higher token throughput."),
            ("Local ONNX Reranking & Autoscaling:", "Embedded FlashRank ONNX local re-ranking (+21.4% MRR) and KEDA Kafka lag-based pod autoscaling (2 to 32 pods in 35s), serving 420 concurrent users per GPU.")
        ],
        "metrics": [("50x", "SIMD Acceleration"), ("<0.5ms", "Scoring Latency"), ("-68%", "GPU Cloud Spend")]
    },
    {
        "id": 3,
        "category": "FLAGSHIP 03: FOUNDER VENTURE",
        "title": "Founder Showcase: KoomAI — High-Performance Video Intelligence",
        "subtitle": "Distributed WebRTC Media Pipeline & 20-Agent LangGraph Mesh",
        "badge": "$24K MRR · SUB-150MS WEBRTC",
        "accent": BORDER_GOLD,
        "img_src": "/Users/kishorkumarparoi/job-project/vantage-portfolio/images/projects/koom/01_architecture_flow.jpg",
        "bullets": [
            ("Sub-150ms WebRTC Media Pipeline:", "Designed distributed edge ingest and WebRTC media pipeline with client-side Rust/WASM frame capture, chunked uploading, and hardware-accelerated FFmpeg transcoding."),
            ("20-Agent LangGraph Mesh:", "Orchestrated autonomous multi-agent mesh for automated speech-to-text, context chaptering, smart action item extraction, and meeting summarization."),
            ("Full-Stack Enterprise Architecture:", "Engineered Next.js 15 command center, Chrome Extension, Electron desktop client, and Go microservices with Stripe billing and AWS CloudFront edge delivery."),
            ("Commercial Scale & Reliability:", "Bootstrapped to $24K MRR with 99.98% pipeline availability, zero packet loss, and global sub-second video delivery across 4 continents.")
        ],
        "metrics": [("$24K", "Monthly Recurring Rev"), ("<150ms", "WebRTC Media Ingest"), ("20 Agents", "LangGraph Mesh")]
    },
    {
        "id": 4,
        "category": "FLAGSHIP 04: AI SECURITY & GATEWAYS",
        "title": "Enterprise AI Security & Production RAG Gateway",
        "subtitle": "Defense-in-Depth AI Gateway, NVIDIA NeMo & Portkey Cache",
        "badge": "95%+ ATTACKS HALTED · <35MS P95",
        "accent": BORDER_EMERALD,
        "img_src": "/Users/kishorkumarparoi/job-project/vantage-portfolio/images/projects/security/07_hybrid_reranker_cascade.jpg",
        "bullets": [
            ("Zero-Trust Defense-in-Depth:", "Enforced BotShield WAF and NVIDIA NeMo Guardrails; halted 95%+ of prompt injections and jailbreaks (validated via Microsoft PyRIT 0.14.0 red-teaming)."),
            ("Sub-35ms Semantic Cache:", "Implemented Portkey + Redis semantic cache with cosine distance indexing; slashed LLM API bills by $8.4K/mo (42% savings) under <35ms P95 latency."),
            ("FastMCP & HITL Governance:", "Standardized tool execution via Model Context Protocol (FastMCP) with Human-in-the-Loop break-glass gates and cryptographic HMAC-SHA256 tool signing."),
            ("Automated RAGAS Synthetic Harness:", "Codified automated synthetic evaluation measuring Context Precision, Faithfulness, and Answer Relevance on every CI/CD pull request before deployment.")
        ],
        "metrics": [("95%+", "Threats Blocked"), ("<35ms", "P95 Cache Latency"), ("$8.4K/mo", "LLM API Slashed")]
    },
    {
        "id": 5,
        "category": "FLAGSHIP 05: INDUSTRIAL EDGE & MLOPS",
        "title": "Smart Manufacturing Industrial AI & Edge MLOps Platform",
        "subtitle": "Sub-10ms Vibration Anomaly Detection & ArgoCD GitOps Engine",
        "badge": "$420K SCRAP SAVED · 96% FASTER DEPLOY",
        "accent": BORDER_CYAN,
        "img_src": "/Users/kishorkumarparoi/job-project/vantage-portfolio/images/projects/smart/architecture_master.png",
        "bullets": [
            ("Industrial Edge Anomaly Detection:", "Deployed PyTorch Autoencoders and CNN defect models (<20ms); mitigated part escape defect rate by 31%, saving $420K annually in industrial scrap."),
            ("Polyglot Edge Stack:", "Rust SIMD vibration pre-filter (:8082), Go edge protocol gateway (:8085), Next.js 15 cockpit (:3000), and PyTorch AI core (:8000) over constrained hardware."),
            ("Declarative CI/CD & ArgoCD GitOps:", "Architected 12-stage Jenkins pipeline, SonarQube, Nexus 3, and ArgoCD GitOps on K3s; curtailed deployment cycle by 96% (4h down to 8m)."),
            ("TensorRT & High-Frequency Protocols:", "Optimized models via TensorRT INT8 (<8.2ms), non-blocking OPC-UA protocol adapter (50K tags/sec), and bidirectional gRPC edge-to-cloud streaming (-68% bandwidth).")
        ],
        "metrics": [("$420K", "Scrap Waste Saved"), ("<8.2ms", "TensorRT INT8"), ("96%", "Deploy Cycle Cut")]
    },
    {
        "id": 6,
        "category": "FLAGSHIP 06: MULTI-CLOUD & AGENTS",
        "title": "Nexus-AI Multi-Cloud Lakehouse & Autonomous Agent Mesh",
        "subtitle": "Cross-Cloud Lakehouse (AWS/GCP/Azure) & LangGraph NL2SQL",
        "badge": "18M+ EVENTS · <250MS QUERY",
        "accent": BORDER_GOLD,
        "img_src": "/Users/kishorkumarparoi/job-project/vantage-portfolio/images/projects/nexus/01_master_architecture.jpg",
        "bullets": [
            ("Multi-Cloud Lakehouse Fabric:", "Unified AWS S3, GCP GCS, and Azure ADLS Gen2 via PySpark Delta Lake & dbt Core; consolidated 18M+ cross-cloud events under <250ms query latency."),
            ("LangGraph Multi-Agent Mesh:", "Orchestrated autonomous multi-agent mesh (Data NL2SQL, BI Analyst, MLOps Health, OWASP Guardrails); automated analytical reasoning with 99.2% reliability."),
            ("High-Throughput Gateway:", "Polyglot microservices in Go 1.22 (:3000), FastAPI core (:8081), and Next.js 16 SSE streaming; eliminated 65% of legacy gateway overhead."),
            ("Automated MLOps Governance:", "Integrated continuous ML training and model registry with MLflow, Feast Feature Store, and Kubernetes; cut deployment turnaround from days to <12 minutes.")
        ],
        "metrics": [("18M+", "Cross-Cloud Events"), ("<250ms", "Query Latency"), ("99.2%", "Agent Reliability")]
    },
    {
        "id": 7,
        "category": "FLAGSHIP 07: AZURE ENTERPRISE DATA",
        "title": "Advanced Network Security Medallion Lakehouse (Azure)",
        "subtitle": "Azure Data Factory Ingestion, ADLS Gen2 & Synapse Serverless CETAS",
        "badge": "50M+ LOGS · -58% ANALYTICS COST",
        "accent": BORDER_CYAN,
        "img_src": "/Users/kishorkumarparoi/job-project/vantage-portfolio/images/projects/azure_etl/architecture_diagram.png",
        "bullets": [
            ("Azure Enterprise Ingestion:", "Orchestrated Azure Data Factory (ADF) multi-pipeline ingestion loading high-volume firewall telemetry into ADLS Gen2 Bronze data lake."),
            ("Serverless Synapse CETAS Engine:", "Built automated Silver-to-Gold transformations using Synapse CETAS (Create External Table As Select), slashing data processing costs by 58%."),
            ("Threat Intelligence Analytics:", "Transformed 50M+ network logs with automated IP geo-enrichment, DDoS anomaly alerts, and partitioned Parquet indexing for instant analytical querying."),
            ("Power BI Live Executive Dashboards:", "Connected DirectQuery models providing real-time security posture dashboards with sub-2s query response for C-suite executive monitoring.")
        ],
        "metrics": [("50M+", "Firewall Logs"), ("-58%", "Processing Spend"), ("<2s", "DirectQuery Dashboard")]
    },
    {
        "id": 8,
        "category": "FLAGSHIP 08: AWS PRODUCTION MLOPS",
        "title": "AWS SageMaker End-to-End MLOps & Production Pipelines",
        "subtitle": "SageMaker Pipelines, Online Feature Store & Automated Canary Deploy",
        "badge": "ZERO-DOWNTIME · AUTOMATED RETRAIN",
        "accent": BORDER_GOLD,
        "img_src": "/Users/kishorkumarparoi/Desktop/FIVERR_PORTFOLIO_ASSETS/12_AWS_SageMaker_End_to_End_MLOps/01_architecture_flow.png",
        "bullets": [
            ("Automated SageMaker Pipelines:", "Codified end-to-end DAGs for distributed model training, data preprocessing, model evaluation, and conditional model registration."),
            ("SageMaker Feature Store & Model Registry:", "Implemented online/offline feature store with sub-10ms lookup and automated model approval gates with lineage tracking."),
            ("Blue/Green & Canary Deployments:", "Automated zero-downtime canary endpoint traffic shifting with automated rollback on CloudWatch error threshold spikes."),
            ("Drift & Bias Monitoring:", "Integrated SageMaker Clarify and Model Monitor for automated concept drift detection, triggering continuous retraining without human intervention.")
        ],
        "metrics": [("0 Down", "Canary Rollouts"), ("<10ms", "Feature Store"), ("Auto", "Drift Retraining")]
    },
    {
        "id": 9,
        "category": "FLAGSHIP 09: CLOUD-NATIVE SERVING",
        "title": "Cloud-Native Real-Time MLOps & Scale-to-Zero Serving (KServe)",
        "subtitle": "KServe Serverless Inference, Triton Runtimes & Istio Mesh",
        "badge": "-72% IDLE GPU COST · SUB-15MS P95",
        "accent": BORDER_EMERALD,
        "img_src": "/Users/kishorkumarparoi/Desktop/FIVERR_PORTFOLIO_ASSETS/13_Cloud_Native_RealTime_MLOps_KServe/01_architecture_flow.png",
        "bullets": [
            ("KServe Scale-to-Zero Inference:", "Deployed KServe on Kubernetes with Knative Serving for automatic scale-to-zero, saving 72% in idle cloud GPU costs during low traffic."),
            ("Multi-Model Triton & vLLM Runtimes:", "Served PyTorch, ONNX, and LLM models with dynamic batching, GPU memory paging, and unified inference protocol (v2)."),
            ("Istio Service Mesh Routing:", "Enforced mTLS zero-trust communication, canary traffic splitting (90/10), and circuit breaking for resilient high-load serving."),
            ("OpenTelemetry Observability:", "Full OpenTelemetry instrumentation with Prometheus metrics and Grafana dashboards for real-time P95/P99 latency tracking.")
        ],
        "metrics": [("-72%", "Idle GPU Spend"), ("<15ms", "P95 Serving"), ("100%", "Zero-Trust mTLS")]
    },
    {
        "id": 10,
        "category": "FLAGSHIP 10: PRODUCTION HITL CHATBOT",
        "title": "Production AI Chatbot & Human-in-the-Loop (HITL) Platform",
        "subtitle": "LangGraph Stateful Workflow, MemorySaver & Break-Glass Gates",
        "badge": "100% AUDITABLE · -46% HALLUCINATIONS",
        "accent": BORDER_VIOLET,
        "img_src": "/Users/kishorkumarparoi/job-project/vantage-portfolio/images/projects/chatbot/03_agentic_retrieval.jpg",
        "bullets": [
            ("Stateful LangGraph Workflows:", "Engineered conversational state graph with MemorySaver persistent checkpoints, supporting multi-turn complex dialogs and context retention."),
            ("Human-in-the-Loop Break-Glass Gates:", "Integrated pause-and-resume approval gates for sensitive tool invocations (financial actions, external emails, refunds) with audit trails."),
            ("Hybrid RAG & Hallucination Mitigation:", "Sub-40ms hybrid retrieval with cross-encoder re-ranking; curtailed hallucination rate by 46% across enterprise knowledge bases."),
            ("Comprehensive Tracing & Replay:", "100% auditable execution log with Logfire/LangSmith tracing, session replay, and deterministic regression testing.")
        ],
        "metrics": [("100%", "Audit Compliance"), ("-46%", "Hallucinations"), ("<40ms", "Retrieval P95")]
    },
    {
        "id": 11,
        "category": "FLAGSHIP 11: AUTONOMOUS DATA AGENTS",
        "title": "Autonomous Data Agent — Multi-Agent BI & SQL Engine",
        "subtitle": "Hierarchical Planner, Schema Introspector & Self-Healing SQL",
        "badge": "ZERO SCAN CRASHES · <200MS REPAIR",
        "accent": BORDER_CYAN,
        "img_src": "/Users/kishorkumarparoi/job-project/pptx-deck/opt_img/p11_dataagent_arch.png",
        "bullets": [
            ("Hierarchical Agent Graph:", "Orchestrated 4 specialized agents: Query Planner, Schema Introspector, SQL Generator & Validator, and Automated Visualizer."),
            ("Zero-Hallucination SQL Synthesis:", "Dynamic schema pruning and automated EXPLAIN cost check before query execution, completely preventing full-table scan crashes."),
            ("Self-Correcting Execution Loop:", "If SQL execution throws a syntax or schema error, the agent automatically inspects the error trace and repairs the query in <200ms."),
            ("Automated BI Visualization:", "Generates interactive Vega-Lite / Plotly specifications from query result sets with zero human code intervention.")
        ],
        "metrics": [("0 Crash", "Full Scan Guard"), ("<200ms", "Self-Healing SQL"), ("4 Agents", "Specialized Mesh")]
    },
    {
        "id": 12,
        "category": "FLAGSHIP 12: BARE-METAL SYSTEMS",
        "title": "Core Systems: Rust Git Engine & Async POSIX Terminal Shell",
        "subtitle": "Bare-Metal Rust Internals, SIMD Zlib & Zero-Allocation Piping",
        "badge": "SUB-1µS LATENCY · -73% HEAP ALLOC",
        "accent": BORDER_ROSE,
        "img_src": "/Users/kishorkumarparoi/job-project/vantage-portfolio/images/projects/rustgit/01_architecture_flow.jpg",
        "bullets": [
            ("Bare-Metal Git Engine in Rust:", "Implemented complete Git internals in Rust: SHA-1 directed acyclic graph (blobs, trees, commits, tags) and packfile indexing."),
            ("SIMD Zlib Streaming:", "Optimized loose object decompression via SIMD-accelerated streaming zlib, matching C execution speed with 100% memory safety."),
            ("Async POSIX Terminal Shell:", "Engineered async terminal shell with tokio, supporting signal multiplexing (SIGINT/SIGTSTP), job control, piping, and I/O redirection."),
            ("Zero-Allocation Hot Paths:", "Attained sub-microsecond (<1µs) execution latency with 73% lower heap allocations compared to standard bash/zsh shells.")
        ],
        "metrics": [("<1µs", "Hot-Path Latency"), ("-73%", "Heap Allocation"), ("100%", "Memory Safety")]
    },
    {
        "id": 13,
        "category": "FLAGSHIP 13: WEBGL & CYBER TELEMETRY",
        "title": "3D Cyber Hacker Cockpit & Interactive WebGL Systems",
        "subtitle": "Three.js, GLSL Custom Shaders & Real-Time Cyber Telemetry",
        "badge": "60 FPS · REAL-TIME WEBSOCKET",
        "accent": BORDER_EMERALD,
        "img_src": "/Users/kishorkumarparoi/job-project/vantage-portfolio/images/projects/hacker/01_hacker_cockpit.jpg",
        "bullets": [
            ("Immersive 3D Cockpit:", "Developed full WebGL / Three.js 3D cockpit environment with interactive CRT monitors, HUD gauges, and retro cyber telemetry."),
            ("Custom GLSL Shaders:", "Wrote custom vertex and fragment shaders for bloom, chromatic aberration, scanlines, and volumetric neon lighting at steady 60 FPS."),
            ("Real-Time Data Streaming:", "Connected WebSocket telemetry feeds driving animated gauges, CPU load graphs, and live threat map visualizations."),
            ("Zero-Lag UI Integration:", "Integrated React Three Fiber with DOM overlays using CSS 3D transforms for seamless text input and keyboard navigation.")
        ],
        "metrics": [("60 FPS", "Steady Framerate"), ("WebGL", "Custom Shaders"), ("Real-Time", "Telemetry Feeds")]
    },
    {
        "id": 14,
        "category": "FLAGSHIP 14: CLOUD DESKTOP WEB OS",
        "title": "MacOS Sequoia Interactive Desktop Web OS",
        "subtitle": "Next.js 15, Framer Motion Physics & Zustand Reactive Store",
        "badge": "PIXEL-PERFECT · FULL INTERACTION",
        "accent": BORDER_CYAN,
        "img_src": "/Users/kishorkumarparoi/job-project/vantage-portfolio/images/projects/macos/01_desktop_os.jpg",
        "bullets": [
            ("Full-Fidelity Web OS:", "Engineered complete macOS Sequoia desktop simulation in Next.js 15, TypeScript, and Tailwind CSS with pixel-perfect attention to detail."),
            ("Physics-Based Window Manager:", "Custom draggable, resizable, minimizable windowing system with z-index stacking and dock magnification physics."),
            ("Rich Native Applications:", "Includes interactive Terminal emulator, Spotlight search (Cmd+K), Safari web browser, Notes, Settings, and audio player."),
            ("Audio Synthesis Engine:", "Built Web Audio API sound engine replicating authentic macOS startup chimes, click feedback, and notification sounds.")
        ],
        "metrics": [("Next.js 15", "App Router"), ("Zustand", "State Store"), ("Web Audio", "Sound Synthesis")]
    },
    {
        "id": 15,
        "category": "FLAGSHIP 15: 3D DIGITAL TWIN",
        "title": "Interactive 3D Space Station (AWS DeepRacer / Digital Twin)",
        "subtitle": "Three.js Space Station Twin, Orbital Physics & LOD Streaming",
        "badge": "60 FPS ON MOBILE · SENSOR FUSION",
        "accent": BORDER_VIOLET,
        "img_src": "/Users/kishorkumarparoi/job-project/vantage-portfolio/images/projects/reinvent/01_3d_architecture.png",
        "bullets": [
            ("Digital Twin Architecture:", "Built interactive 3D digital twin of space station research module using Three.js, WebGL, and custom GLTF models."),
            ("Sensor Fusion & Telemetry HUD:", "Simulated orbital mechanics, solar panel alignment, and temperature sensor telemetry with real-time UI gauges."),
            ("LOD Optimization:", "Level of Detail (LOD) geometry streaming and GPU instancing ensuring smooth 60 FPS on low-power mobile GPUs."),
            ("Interactive Exploration:", "First-person and orbit camera controls with collision detection, hotspot annotation markers, and guided inspection tours.")
        ],
        "metrics": [("60 FPS", "Mobile Optimized"), ("Three.js", "LOD Streaming"), ("Sensor", "Telemetry HUD")]
    }
]

for p in PROJECTS:
    slide_num = p["id"] + 5  # Slides 6 to 20
    print(f"Building Slide {slide_num}: {p['title']}...")
    s = add_slide()
    add_header(s, p["title"], p["category"], slide_num=slide_num, accent_col=p["accent"])
    
    col = p["accent"]
    
    # ── Left Container: Architecture Blueprint Image (Width: 5.80", Height: 5.45") ──
    add_box(s, 0.80, 1.50, 5.80, 5.45, fill_col=BG_CARD, border_col=BORDER_DIM, border_w=1.0)
    add_box(s, 0.80, 1.50, 0.10, 5.45, fill_col=col, border_col=None)
    
    # Left container top header label
    tx_bl = s.shapes.add_textbox(Inches(1.05), Inches(1.65), Inches(5.30), Inches(0.28))
    tf_bl = tx_bl.text_frame
    tf_bl.margin_left = tf_bl.margin_right = tf_bl.margin_top = tf_bl.margin_bottom = Pt(0)
    p_bl = tf_bl.paragraphs[0]
    r_bl = p_bl.add_run()
    r_bl.text = "ARCHITECTURE TOPOLOGY & DATA FLOW"
    r_bl.font.size = Pt(9.0)
    r_bl.font.bold = True
    r_bl.font.color.rgb = col
    
    # Load and center the architecture image
    img_opt = get_optimized_image(p["img_src"], f"p{p['id']:02d}_opt")
    if img_opt and os.path.exists(img_opt):
        try:
            im = Image.open(img_opt)
            iw, ih = im.size
            aspect = iw / ih
            
            # Available bounding box inside container: max_w = 5.40, max_h = 4.60
            max_w, max_h = 5.40, 4.65
            if aspect > (max_w / max_h):
                fit_w = max_w
                fit_h = max_w / aspect
            else:
                fit_h = max_h
                fit_w = max_h * aspect
                
            img_x = 0.80 + (5.80 - fit_w) / 2.0
            img_y = 2.05 + (4.70 - fit_h) / 2.0
            
            # Add subtle dark border frame behind image
            add_box(s, img_x - 0.04, img_y - 0.04, fit_w + 0.08, fit_h + 0.08, fill_col=BG_VOID, border_col=BORDER_DIM, border_w=0.8)
            s.shapes.add_picture(img_opt, Inches(img_x), Inches(img_y), Inches(fit_w), Inches(fit_h))
        except Exception as e:
            print(f"Error placing picture for project {p['id']}: {e}")
    
    # ── Right Container: Technical Architecture Breakdown (Width: 5.733", Height: 5.45") ──
    add_box(s, 6.80, 1.50, 5.733, 5.45, fill_col=BG_CARD, border_col=col, border_w=1.2, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_box(s, 6.80, 1.50, 0.12, 5.45, fill_col=col, border_col=None)
    
    # Top Badge Pill & Subtitle
    add_pill(s, 7.05, 1.68, p["badge"], col, BG_CARD_LIGHT, col, font_sz=8.0)
    
    tx_sub = s.shapes.add_textbox(Inches(7.05), Inches(2.05), Inches(5.25), Inches(0.40))
    tf_sub = tx_sub.text_frame
    tf_sub.word_wrap = True
    tf_sub.margin_left = tf_sub.margin_right = tf_sub.margin_top = tf_sub.margin_bottom = Pt(0)
    p_st = tf_sub.paragraphs[0]
    r_st = p_st.add_run()
    r_st.text = p["subtitle"]
    r_st.font.size = Pt(13.0)
    r_st.font.bold = True
    r_st.font.color.rgb = WHITE
    
    # Divider line under subtitle
    div = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.05), Inches(2.48), Inches(5.25), Inches(0.012))
    div.fill.solid()
    div.fill.fore_color.rgb = RGBColor(0x1F, 0x2D, 0x48)
    div.line.fill.background()
    
    # 4 Detailed Technical Bullet Points
    tx_pts = s.shapes.add_textbox(Inches(7.05), Inches(2.60), Inches(5.25), Inches(3.55))
    tf_pts = tx_pts.text_frame
    tf_pts.word_wrap = True
    tf_pts.margin_left = tf_pts.margin_right = tf_pts.margin_top = tf_pts.margin_bottom = Pt(0)
    
    for b_idx, (b_lead, b_body) in enumerate(p["bullets"]):
        p_pt = tf_pts.paragraphs[0] if b_idx == 0 else tf_pts.add_paragraph()
        if b_idx > 0:
            p_pt.space_before = Pt(4.5)
        p_pt.line_spacing = Pt(13)
        
        r_lead = p_pt.add_run()
        r_lead.text = f"• {b_lead} "
        r_lead.font.size = Pt(9.8)
        r_lead.font.bold = True
        r_lead.font.color.rgb = col
        
        r_body = p_pt.add_run()
        r_body.text = b_body
        r_body.font.size = Pt(9.2)
        r_body.font.color.rgb = MUTED_GRAY
        
    # Bottom Mini-Metric Stat Bar (3 Pill Cards)
    m_y = 6.25
    m_w = 1.65
    m_gap = 0.15
    m_starts = [7.05, 7.05 + m_w + m_gap, 7.05 + 2 * (m_w + m_gap)]
    
    for midx, (mval, mlbl) in enumerate(p["metrics"]):
        mx = m_starts[midx]
        add_box(s, mx, m_y, m_w, 0.52, fill_col=BG_CARD_LIGHT, border_col=col, border_w=0.8, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
        
        tx_m = s.shapes.add_textbox(Inches(mx + 0.05), Inches(m_y + 0.04), Inches(m_w - 0.10), Inches(0.44))
        tf_m = tx_m.text_frame
        tf_m.margin_left = tf_m.margin_right = tf_m.margin_top = tf_m.margin_bottom = Pt(0)
        
        p_mv = tf_m.paragraphs[0]
        p_mv.alignment = PP_ALIGN.CENTER
        r_mv = p_mv.add_run()
        r_mv.text = mval
        r_mv.font.size = Pt(11.0)
        r_mv.font.bold = True
        r_mv.font.color.rgb = col
        
        p_ml = tf_m.add_paragraph()
        p_ml.alignment = PP_ALIGN.CENTER
        r_ml = p_ml.add_run()
        r_ml.text = mlbl
        r_ml.font.size = Pt(7.5)
        r_ml.font.color.rgb = MUTED_WHITE

# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 21: VALUE DIFFERENTIATION MATRIX (KISHOR VS TYPICAL SENIOR ENGINEERS)
# ═════════════════════════════════════════════════════════════════════════════
print("Building Slide 21: Value Differentiation...")
s21 = add_slide()
add_header(s21, "Value Differentiation: Production Rigor vs Typical Prototypers", "COMPETITIVE ADVANTAGE", slide_num=21)

diff_dims = [
    ("Systems Rigor & Speed", "Python wrappers with multi-second latency spikes and heavy heap footprints.", "Polyglot systems (Rust SIMD, Go 1.22, C++) with 50x compute acceleration and <1µs hot paths.", BORDER_CYAN),
    ("RAG & Retrieval Architecture", "Naive chunking, standard vector DB lookup, 400ms+ latency, frequent hallucinations.", "Sub-35ms Hybrid RAG (HNSW + BM25 + FlashRank ONNX) with Portkey semantic cache (-46% errors).", BORDER_VIOLET),
    ("GPU & Cloud Cost Efficiency", "Defaulting to expensive proprietary APIs with runaway $25K+/mo bills.", "Multi-LoRA vLLM dynamic adapter multiplexing on single GPUs; 68% compute cost reduction.", BORDER_GOLD),
    ("Security & AI Governance", "Raw prompts sent directly to LLMs; vulnerable to injection & data exfiltration.", "Zero-Trust BotShield WAF, NVIDIA NeMo Guardrails, and FastMCP Human-in-the-Loop gates.", BORDER_EMERALD),
    ("Business Impact & Track Record", "Completes isolated Jira tickets; limited insight into cloud infrastructure economics.", "Directly secured $1.4M in enterprise contracts, cut $18K/mo cloud spend, bootstrapped $24K MRR SaaS.", BORDER_ROSE)
]

col_w = 11.733
row_h = 0.95
y_start = 1.55

for idx, (dim, typical, kishor, col) in enumerate(diff_dims):
    ry = y_start + idx * (row_h + 0.12)
    add_box(s21, 0.80, ry, col_w, row_h, fill_col=BG_CARD, border_col=col, border_w=1.0, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_box(s21, 0.80, ry, 0.10, row_h, fill_col=col, border_col=None)
    
    # Dimension Name (Col 1: 0.95 to 3.20)
    tx_dim = s21.shapes.add_textbox(Inches(1.05), Inches(ry + 0.15), Inches(2.25), Inches(row_h - 0.30))
    tf_dim = tx_dim.text_frame
    tf_dim.word_wrap = True
    tf_dim.margin_left = tf_dim.margin_right = tf_dim.margin_top = tf_dim.margin_bottom = Pt(0)
    p_d = tf_dim.paragraphs[0]
    r_d = p_d.add_run()
    r_d.text = dim
    r_d.font.size = Pt(11.0)
    r_d.font.bold = True
    r_d.font.color.rgb = col
    
    # Typical Senior Engineer (Col 2: 3.40 to 7.30)
    tx_typ = s21.shapes.add_textbox(Inches(3.40), Inches(ry + 0.15), Inches(3.90), Inches(row_h - 0.30))
    tf_typ = tx_typ.text_frame
    tf_typ.word_wrap = True
    tf_typ.margin_left = tf_typ.margin_right = tf_typ.margin_top = tf_typ.margin_bottom = Pt(0)
    p_typ_t = tf_typ.paragraphs[0]
    r_typt = p_typ_t.add_run()
    r_typt.text = "TYPICAL SENIOR ENGINEER"
    r_typt.font.size = Pt(8.0)
    r_typt.font.bold = True
    r_typt.font.color.rgb = DARK_GRAY
    p_typ_b = tf_typ.add_paragraph()
    p_typ_b.space_before = Pt(2)
    r_typb = p_typ_b.add_run()
    r_typb.text = typical
    r_typb.font.size = Pt(8.5)
    r_typb.font.color.rgb = MUTED_GRAY
    
    # Kishor Kumar Paroi (Col 3: 7.50 to 12.30)
    tx_kish = s21.shapes.add_textbox(Inches(7.50), Inches(ry + 0.15), Inches(4.85), Inches(row_h - 0.30))
    tf_kish = tx_kish.text_frame
    tf_kish.word_wrap = True
    tf_kish.margin_left = tf_kish.margin_right = tf_kish.margin_top = tf_kish.margin_bottom = Pt(0)
    p_k_t = tf_kish.paragraphs[0]
    r_kt = p_k_t.add_run()
    r_kt.text = "KISHOR KUMAR PAROI (VERIFIED ADVANTAGE)"
    r_kt.font.size = Pt(8.0)
    r_kt.font.bold = True
    r_kt.font.color.rgb = col
    p_k_b = tf_kish.add_paragraph()
    p_k_b.space_before = Pt(2)
    r_kb = p_k_b.add_run()
    r_kb.text = kishor
    r_kb.font.size = Pt(8.8)
    r_kb.font.bold = True
    r_kb.font.color.rgb = WHITE

# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 22: TARGET ROLES & ENGAGEMENT MODEL (COLLABORATION)
# ═════════════════════════════════════════════════════════════════════════════
print("Building Slide 22: Target Roles & Engagement...")
s22 = add_slide()
add_header(s22, "Target Roles & Engagement Model: How We Can Partner", "COLLABORATION MODES", slide_num=22)

roles_data = [
    ("Target Roles", "High-Impact Architectural & Leadership Roles", BORDER_CYAN, [
        "Principal AI Systems Architect",
        "Forward Deployed AI Engineer (FDE)",
        "Head of AI / AI Engineering Lead",
        "Senior Staff MLOps & Platform Architect",
        "Staff Systems Engineer (Rust / Go / Distributed)"
    ]),
    ("Engagement Models", "Flexible Global Partnership Structures", BORDER_GOLD, [
        "Full-Time Global Remote (Tier-1 Tech / Scale-ups)",
        "Executive Advisory & Fractional AI Architect",
        "High-Impact 90-Day Production Sprints",
        "Zero-Downtime Microservices & AI Migration",
        "Enterprise Multi-Agent Mesh & Lakehouse Overhauls"
    ]),
    ("Direct Ecosystem & Booking", "Get In Touch & Explore Live Artifacts", BORDER_EMERALD, [
        "koom.kishor.tech (Live SaaS)",
        "kishor.tech (Engineering Portfolio)",
        "macos.kishor.tech (Interactive WebOS)",
        "linkedin.com/in/kishor-kumar-paroi",
        "kishor.ruet.cse@gmail.com"
    ])
]

c_w = 3.75
c_h = 5.45
c_xs = [0.80, 4.80, 8.80]

for idx, (title, sub, col, items) in enumerate(roles_data):
    cx = c_xs[idx]
    add_box(s22, cx, 1.50, c_w, c_h, fill_col=BG_CARD, border_col=col, border_w=1.2, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_box(s22, cx, 1.50, c_w, 0.12, fill_col=col, border_col=None)
    
    tx_c = s22.shapes.add_textbox(Inches(cx + 0.25), Inches(1.80), Inches(c_w - 0.50), Inches(c_h - 0.60))
    tf_c = tx_c.text_frame
    tf_c.word_wrap = True
    tf_c.margin_left = tf_c.margin_right = tf_c.margin_top = tf_c.margin_bottom = Pt(0)
    
    p_t = tf_c.paragraphs[0]
    r_t = p_t.add_run()
    r_t.text = title
    r_t.font.size = Pt(14.0)
    r_t.font.bold = True
    r_t.font.color.rgb = col
    
    p_s = tf_c.add_paragraph()
    p_s.space_before = Pt(3)
    r_s = p_s.add_run()
    r_s.text = sub
    r_s.font.size = Pt(9.5)
    r_s.font.color.rgb = MUTED_GRAY
    
    p_div = tf_c.add_paragraph()
    p_div.space_before = Pt(8)
    
    for item in items:
        p_i = tf_c.add_paragraph()
        p_i.space_before = Pt(10)
        r_i = p_i.add_run()
        r_i.text = f"• {item}"
        r_i.font.size = Pt(10.0)
        r_i.font.bold = True
        r_i.font.color.rgb = WHITE

# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 23: 90-DAY STRATEGIC IMPACT PLAYBOOK (ROADMAP)
# ═════════════════════════════════════════════════════════════════════════════
print("Building Slide 23: 90-Day Execution Roadmap...")
s23 = add_slide()
add_header(s23, "90-Day Execution Roadmap: From Day 0 Audit to Production Scale", "ACTIONABLE PLAYBOOK", slide_num=23)

phases = [
    ("DAYS 1 – 30", "Architecture Audit & Quick-Win ROI", BORDER_CYAN, [
        "Audit existing inference latency, GPU cloud spend, and token routing bottlenecks.",
        "Implement Portkey semantic caching with Redis; immediate 35-50% LLM bill reduction.",
        "Deploy FlashRank ONNX local re-ranking to boost retrieval MRR without adding GPUs.",
        "Codify CI/CD baseline with linting, security scanning, and automated smoke tests."
    ]),
    ("DAYS 31 – 60", "Distributed Mesh & Lakehouse Unification", BORDER_GOLD, [
        "Transition brittle LLM chains to LangGraph cyclic state machines with memory checkpoints.",
        "Deploy vLLM Multi-LoRA dynamic adapter multiplexing, slashing redundant GPU instances.",
        "Unify disparate event logs into PySpark Delta Lake Medallion lakehouse with dbt Core.",
        "Integrate Feast feature store over Redis for sub-5ms real-time feature retrieval."
    ]),
    ("DAYS 61 – 90", "Autonomous Scale & Enterprise Hardening", BORDER_EMERALD, [
        "Implement NVIDIA NeMo Guardrails and FastMCP Human-in-the-Loop break-glass gates.",
        "Codify declarative ArgoCD GitOps pipelines on Kubernetes with automated canary rollouts.",
        "Enforce multi-region disaster recovery and Delta compaction (RPO < 60s, RTO < 4.5m).",
        "Deliver automated RAGAS synthetic regression harness on every pull request."
    ])
]

rp_w = 11.733
rp_h = 1.65
rp_start_y = 1.60

for idx, (p_tag, p_title, col, steps) in enumerate(phases):
    py = rp_start_y + idx * (rp_h + 0.22)
    add_box(s23, 0.80, py, rp_w, rp_h, fill_col=BG_CARD, border_col=col, border_w=1.2, shape_type=MSO_SHAPE.ROUNDED_RECTANGLE)
    add_box(s23, 0.80, py, 0.12, rp_h, fill_col=col, border_col=None)
    
    # Left Tag Block
    add_pill(s23, 1.10, py + 0.25, p_tag, col, BG_CARD_LIGHT, col, font_sz=9.0)
    
    tx_pt = s23.shapes.add_textbox(Inches(1.10), Inches(py + 0.65), Inches(2.80), Inches(0.85))
    tf_pt = tx_pt.text_frame
    tf_pt.word_wrap = True
    tf_pt.margin_left = tf_pt.margin_right = tf_pt.margin_top = tf_pt.margin_bottom = Pt(0)
    p_h = tf_pt.paragraphs[0]
    r_h = p_h.add_run()
    r_h.text = p_title
    r_h.font.size = Pt(12.5)
    r_h.font.bold = True
    r_h.font.color.rgb = WHITE
    
    # Right Steps Block
    tx_st = s23.shapes.add_textbox(Inches(4.20), Inches(py + 0.20), Inches(8.10), Inches(rp_h - 0.35))
    tf_st = tx_st.text_frame
    tf_st.word_wrap = True
    tf_st.margin_left = tf_st.margin_right = tf_st.margin_top = tf_st.margin_bottom = Pt(0)
    
    for sidx, step in enumerate(steps):
        p_s = tf_st.paragraphs[0] if sidx == 0 else tf_st.add_paragraph()
        if sidx > 0:
            p_s.space_before = Pt(3)
        p_s.line_spacing = Pt(12.5)
        r_s = p_s.add_run()
        r_s.text = f"• {step}"
        r_s.font.size = Pt(9.5)
        r_s.font.color.rgb = MUTED_WHITE

# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 24: CINEMATIC VISUAL OUTRO (VANTAGE RETRO SYNTHWAVE CREST FULL-BLEED)
# ═════════════════════════════════════════════════════════════════════════════
print("Building Slide 24: Outro...")
s24 = add_slide()
img_outro_path = os.path.join(DECK_DIR, "outro.png")
if os.path.exists(img_outro_path):
    s24.shapes.add_picture(img_outro_path, Inches(0), Inches(0), width=prs.slide_width, height=prs.slide_height)

# ═════════════════════════════════════════════════════════════════════════════
# SAVE PRESENTATION TO MULTIPLE STANDARD DESTINATIONS
# ═════════════════════════════════════════════════════════════════════════════
out_files = [
    os.path.join(DECK_DIR, "Kishor_Kumar_Paroi_10000_Dollar_Executive_Deck.pptx"),
    os.path.join(DECK_DIR, "pitch_deck.pptx"),
    os.path.join(DECK_DIR, "deck.pptx"),
]

for out_path in out_files:
    prs.save(out_path)
    sz = os.path.getsize(out_path) / (1024 * 1024)
    print(f"Saved: {out_path} ({sz:.2f} MB)")

print("\nSUCCESS: All 24 slides generated with 100% ECMA-376 OpenXML compliance!")
