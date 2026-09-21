import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def create_resume():
    doc = docx.Document()
    
    # Page setup - 0.5 inch margins for clean 2-page executive format
    for section in doc.sections:
        section.top_margin = Inches(0.48)
        section.bottom_margin = Inches(0.48)
        section.left_margin = Inches(0.55)
        section.right_margin = Inches(0.55)
        
    # Color palette
    navy = RGBColor(12, 35, 64)        # Deep Navy / Header
    slate_teal = RGBColor(0, 95, 75)   # GE Vernova Green / Subtitles
    charcoal = RGBColor(35, 35, 35)    # Body text
    dark_gray = RGBColor(70, 70, 70)   # Metadata / Dates
    
    def add_bottom_border(paragraph):
        pPr = paragraph._p.get_or_add_pPr()
        pBdr = parse_xml(r'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                         r'<w:bottom w:val="single" w:sz="6" w:space="2" w:color="0C2340"/>'
                         r'</w:pBdr>')
        pPr.append(pBdr)

    # Name Header
    p_name = doc.add_paragraph()
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(1)
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_name = p_name.add_run("MICHAEL THOMPSON, M.S.")
    run_name.font.name = "Arial"
    run_name.font.size = Pt(18)
    run_name.font.bold = True
    run_name.font.color.rgb = navy

    # Target Title Banner
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("SENIOR DIRECTOR — ENTERPRISE OT CYBERSECURITY & OPERATIONAL RESILIENCE")
    run_title.font.name = "Arial"
    run_title.font.size = Pt(10.5)
    run_title.font.bold = True
    run_title.font.color.rgb = slate_teal

    # Sub-banner
    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(3)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run("Global Manufacturing Operations | Purdue Model Architecture | Connected Labs | Co-Author, NIST SP 800-82 Rev 3")
    run_sub.font.name = "Arial"
    run_sub.font.size = Pt(9)
    run_sub.font.bold = True
    run_sub.font.color.rgb = navy

    # Contact line
    p_contact = doc.add_paragraph()
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(6)
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_contact = p_contact.add_run("Baltimore, MD  |  (410) 530-4439  |  Mthompson6782@gmail.com  |  linkedin.com/in/michaelthompson-exec")
    run_contact.font.name = "Arial"
    run_contact.font.size = Pt(8.5)
    run_contact.font.color.rgb = dark_gray

    def add_section_header(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title.upper())
        run.font.name = "Arial"
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = navy
        add_bottom_border(p)
        return p

    def add_bullet(text, category=None):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(0.5)
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.line_spacing = 1.05
        if category:
            run_cat = p.add_run(category + ": ")
            run_cat.font.name = "Arial"
            run_cat.font.size = Pt(8.5)
            run_cat.font.bold = True
            run_cat.font.color.rgb = navy
        run_txt = p.add_run(text)
        run_txt.font.name = "Arial"
        run_txt.font.size = Pt(8.5)
        run_txt.font.color.rgb = charcoal
        return p

    # 1. Executive Summary
    add_section_header("Executive Summary")
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(2)
    p_sum.paragraph_format.space_after = Pt(4)
    p_sum.paragraph_format.line_spacing = 1.08
    run_sum = p_sum.add_run(
        "Accomplished engineering executive and Johns Hopkins-trained Systems Engineer with 25+ years directing enterprise "
        "Operational Technology (OT) cybersecurity, industrial automation, and critical infrastructure resilience across multi-business unit "
        "global conglomerates. Co-author of NIST SP 800-82 Rev 3 (Guide to Operational Technology Security) and active committee "
        "contributor for IEC 62443 / ISA 99 and ISA 84. Proven record defining enterprise OT security strategies, governance frameworks, "
        "and multi-year roadmaps across global manufacturing sites, production facilities, and connected engineering test labs.\n\n"
        "Trusted executive advisor partnering across CISO, CTO, and Business Unit DT Global Supply Chain organizations to balance "
        "Safety, Quality, Delivery, and Cost (SQDC). Expert in operationalizing Purdue Model network segmentation (Levels 0–5, Level 3.5 OT DMZs), "
        "deploying OT visibility platforms integrated with ServiceNow CMDB, establishing Enterprise Vulnerability Management (EVM) "
        "standard work with compensating controls for legacy assets, and enforcing robust remote access (Jumpbox, MFA, HPA). Proven "
        "leader of global engineering direct reports with multi-million-dollar CAPEX/OPEX accountability, translating complex cyber-physical "
        "risk into prioritized investment decisions for the Executive Leadership Team (ELT) and Board of Directors."
    )
    run_sum.font.name = "Arial"
    run_sum.font.size = Pt(8.5)
    run_sum.font.color.rgb = charcoal

    # 2. Executive Core Competencies
    add_section_header("Executive Core Competencies")
    add_bullet(
        "Enterprise OT Security Strategy & Multi-Year Roadmaps • NIST SP 800-82 Rev 3 & IEC 62443 Standards • "
        "Minimum Security Requirements & Policy Authoring • Executive Leadership Team (ELT) & Board Advisory • "
        "CAPEX/OPEX Budget Prioritization • SQDC Operational Risk Balancing",
        "Strategy & Governance"
    )
    add_bullet(
        "Purdue Model (Levels 0–5) & Level 3.5 DMZ Architecture • Next-Gen Firewalls (NGFW) & VLAN Micro-Segmentation • "
        "OT Asset Visibility & ServiceNow CMDB Integration • Enterprise Vulnerability Management (EVM) Standard Work • "
        "Remote Access: Jumpbox, MFA & High Privileged Access (HPA) • Compensating Controls for Unpatchable Legacy OT",
        "Architecture & Delivery"
    )
    add_bullet(
        "Global OT Cybersecurity Team Leadership (Direct Reports) • Matrixed Coalition Building (CISO / CTO / Supply Chain DT) • "
        "Connected Lab & Testbed Security Formalization • Detection & Telemetry (Splunk, CrowdStrike, Qualys) • "
        "Incident Response & Operational Continuity • Lean & Six Sigma Operational Discipline (Master Black Belt)",
        "Leadership & Operations"
    )

    # 3. Professional Experience
    add_section_header("Professional Experience")

    def add_job_header(company, location, title, dates, context=None):
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        
        # Set widths
        table.columns[0].width = Inches(5.1)
        table.columns[1].width = Inches(2.3)
        
        cell_left = table.cell(0, 0)
        cell_right = table.cell(0, 1)
        
        # Clear padding
        for cell in (cell_left, cell_right):
            tcPr = cell._tc.get_or_add_tcPr()
            tcMar = OxmlElement('w:tcMar')
            for m in ('top', 'bottom', 'left', 'right'):
                node = OxmlElement(f'w:{m}')
                node.set(qn('w:w'), '0')
                node.set(qn('w:type'), 'dxa')
                tcMar.append(node)
            tcPr.append(tcMar)
            
        p_l = cell_left.paragraphs[0]
        p_l.paragraph_format.space_before = Pt(4)
        p_l.paragraph_format.space_after = Pt(0)
        p_l.paragraph_format.keep_with_next = True
        
        run_comp = p_l.add_run(company)
        run_comp.font.name = "Arial"
        run_comp.font.size = Pt(9.5)
        run_comp.font.bold = True
        run_comp.font.color.rgb = navy
        
        run_pipe = p_l.add_run(f" | {location}\n")
        run_pipe.font.name = "Arial"
        run_pipe.font.size = Pt(8.5)
        run_pipe.font.color.rgb = dark_gray
        
        run_t = p_l.add_run(title)
        run_t.font.name = "Arial"
        run_t.font.size = Pt(9)
        run_t.font.bold = True
        run_t.font.color.rgb = slate_teal
        
        p_r = cell_right.paragraphs[0]
        p_r.paragraph_format.space_before = Pt(4)
        p_r.paragraph_format.space_after = Pt(0)
        p_r.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_r.paragraph_format.keep_with_next = True
        
        run_d = p_r.add_run(dates)
        run_d.font.name = "Arial"
        run_d.font.size = Pt(8.5)
        run_d.font.bold = True
        run_d.font.color.rgb = dark_gray

        if context:
            p_ctx = doc.add_paragraph()
            p_ctx.paragraph_format.space_before = Pt(1)
            p_ctx.paragraph_format.space_after = Pt(2)
            p_ctx.paragraph_format.keep_with_next = True
            run_c = p_ctx.add_run(context)
            run_c.font.name = "Arial"
            run_c.font.size = Pt(8)
            run_c.font.italic = True
            run_c.font.color.rgb = dark_gray

    # MITRE
    add_job_header(
        "THE MITRE CORPORATION", "Critical Infrastructure & ICS Cybersecurity",
        "Principal Engineer & Outcome Lead (Supporting NIST, DOE, CISA)", "2021 – Present",
        "Direct national-scale cyber-physical security programs, author foundational federal standards, and engineer defense-in-depth protections for critical energy, manufacturing, and nuclear facilities."
    )
    add_bullet(
        "Co-authored the benchmark federal guidance, NIST SP 800-82 Rev 3 (Guide to Operational Technology Security), defining modern ICS defense principles, Purdue Model zone-conduit design, and minimum security baselines implemented across global critical infrastructure.",
        "Federal Standard Co-Authorship"
    )
    add_bullet(
        "Engineered fail-safe digital instrumentation and control (I&C) architectures, safety instrumented system (SIS) protections, and zero-trust OT network boundaries for next-generation nuclear facilities (SMR/MMR) and power generation infrastructure.",
        "Energy & Nuclear Reactor Resilience"
    )
    add_bullet(
        "Formalized cybersecurity architecture and continuous monitoring protocols for national-scale energy testbeds and OT research labs, preventing lateral traversal between enterprise networks and cyber-physical experimentation environments.",
        "Lab & Connected Testbed Security"
    )
    add_bullet(
        "Conducted systemic cyber-physical threat modeling, non-intrusive OT security evaluations, and supply chain vulnerability assessments across energy generation, rail, and chemical manufacturing sectors.",
        "Vulnerability & Threat Modeling"
    )
    add_bullet(
        "Briefed federal agency executives, utility leadership, and congressional committees, translating complex OT vulnerabilities and attack vectors into clear risk mitigation roadmaps and multi-million-dollar capital investment strategies.",
        "Executive & Agency Advisory"
    )

    # CAPGEMINI
    add_job_header(
        "CAPGEMINI", "Technology Consulting",
        "Senior Consultant / Program Manager — OT Technology & Industrial Systems", "2019 – 2021",
        "Directed cybersecurity roadmaps and digital modernization programs for heavily regulated energy, utility, and process manufacturing clients."
    )
    add_bullet(
        "Formulated comprehensive multi-year OT security strategies and compliance frameworks aligned with IEC 62443, NIST 800-82, and NERC CIP, eliminating governance gaps between corporate IT and manufacturing plant operations.",
        "Strategic Security Roadmaps"
    )
    add_bullet(
        "Directed engineering teams in designing Purdue Level 3.5 DMZs, deploying Next-Generation Firewalls, and enforcing strict firewall conduit rules to secure legacy DCS and SCADA networks without risking production downtime.",
        "OT Network Segmentation & Governance"
    )
    add_bullet(
        "Managed full-lifecycle security workstreams, leading cross-functional consultant and client teams, tracking program KPIs, and delivering milestone progress reviews to client executive leadership.",
        "Program Delivery & Execution"
    )

    # AWC
    add_job_header(
        "AWC, INC.", "Industrial Automation & Technology Solutions",
        "Chief Engineer and Technology Officer — OT Technology", "2018 – 2019",
        "Steered technology strategy, automation engineering, and OT cybersecurity policy across industrial manufacturing and municipal infrastructure accounts."
    )
    add_bullet(
        "Authored enterprise-wide IT/OT cybersecurity governance standards and functional safety policies adhering to ANSI/ISA 84 / IEC 61508 and ANSI/ISA 99 / IEC 62443.",
        "IT/OT Cyber & Safety Governance"
    )
    add_bullet(
        "Integrated advanced process control, robotic automation, and industrial IoT telemetry, driving a 40% reduction in operating costs while elevating operational uptime and asset visibility.",
        "Operational Efficiency & Automation"
    )
    add_bullet(
        "Established standardized defense-in-depth compensating controls (network isolation, unidirectional data diodes, read-only gateways) for unpatchable field controllers.",
        "Compensating Controls for Legacy Assets"
    )

    # KOCH INDUSTRIES
    add_job_header(
        "KOCH INDUSTRIES", "Global Multi-Segment Manufacturing Conglomerate",
        "Director of Business Continuity & ICS Cyber Resilience", "2016 – 2018",
        "Directed enterprise-wide industrial control system cybersecurity, operational resilience, and disaster recovery across a diversified global manufacturing footprint spanning refining, chemicals, pulp/paper, and materials."
    )
    add_bullet(
        "Established unified OT security governance, minimum security baselines, and vulnerability management standards across dozens of global manufacturing facilities operating under distinct business unit leadership.",
        "Enterprise OT Security Governance"
    )
    add_bullet(
        "Partnered closely with C-level executives, plant managers, and corporate IT/DT leadership to align cybersecurity investments with business unit production targets, balancing safety, quality, delivery, and cost (SQDC).",
        "Multi-Business Unit Alignment"
    )
    add_bullet(
        "Architected secure vendor remote access frameworks utilizing jumpbox bastions, mandatory multi-factor authentication (MFA), and session recording, eliminating unmonitored modem and third-party VPN entry points.",
        "Remote Access & HPA Controls"
    )
    add_bullet(
        "Built, led, and mentored a high-performing global team of 20+ direct engineers and analysts; instituted training programs that upskilled over 100 plant floor engineers in cyber-physical resilience.",
        "Global Team Leadership & Mentorship"
    )
    add_bullet(
        "Authored enterprise cyber-physical incident response playbooks and executed simulated crisis response exercises, establishing rapid containment and fail-safe recovery procedures for manufacturing lines.",
        "Incident Response & Operational Playbooks"
    )

    # KOCH INDUSTRIES - OPEX
    add_job_header(
        "KOCH INDUSTRIES", "Manufacturing Operations",
        "Operational Excellence Automation Engineering Leader", "2015 – 2016",
        "Oversaw automation architecture, process modernization, and operational excellence for large-scale industrial assets."
    )
    add_bullet(
        "Guided automation engineering and technology integration across capital modernization programs valued at $1.7 billion, ensuring secure-by-design principles were embedded into new plant commissioning.",
        "$1.7B Capital Program Stewardship"
    )
    add_bullet(
        "Applied Six Sigma Master Black Belt methodologies to process control systems, driving multimillion-dollar operational cost efficiencies, reducing process variability, and minimizing unplanned downtime.",
        "Lean & Six Sigma Excellence"
    )

    # SPARTAN TECHNOLOGY
    add_job_header(
        "SPARTAN TECHNOLOGY", "Industrial Technology Consulting",
        "Founder & Owner", "2013 – 2015",
        "Built and scaled an elite OT engineering and cybersecurity consulting practice delivering turnkey automation and network security architectures."
    )
    add_bullet(
        "Scaled consultancy from inception to $1M in net profit within two years, delivering high-stakes automation, SCADA modernization, and cybersecurity hardening for tier-1 industrial clients.",
        "Commercial Scale & Leadership"
    )
    add_bullet(
        "Led multidisciplinary engineering teams in deploying hardened PLC/DCS architectures, secure wireless sensor networks, and Purdue Level 2/3 network segmentation.",
        "Turnkey OT Deployments"
    )

    # EXELON
    add_job_header(
        "EXELON (NUCLEAR DIVISION)", "Power Generation & High-Reliability Operations",
        "Senior Cyber Engineer / Instrumentation & Control (I&C) Engineer 3", "2009 – 2013",
        "Led digital instrumentation & control (I&C) engineering and cybersecurity compliance across nuclear power generation facilities."
    )
    add_bullet(
        "Enforced stringent cybersecurity architectures under NRC 10 CFR 73.54 across safety-critical nuclear digital control networks, establishing uncompromising zero-fail engineering disciplines.",
        "High-Reliability Compliance & Hardening"
    )
    add_bullet(
        "Designed and validated secure digital control loops, protective relay systems, and supervisory interfaces, authoring engineering protocols adopted as industry-wide best practices.",
        "Digital Systems Engineering"
    )

    # TAI ENGINEERING
    add_job_header(
        "TAI ENGINEERING", "Industrial Engineering Services",
        "Instrumentation & Controls (I&C) Division Manager", "2006 – 2009",
        "Led 15+ systems engineers delivering complex automation, validation, and control system integration for regulated life sciences clients."
    )
    add_bullet(
        "Managed division P&L, project staffing, and client delivery; achieved 100% compliance with FDA cGMP / 21 CFR Part 11 electronic records and data integrity standards.",
        "Division Leadership & P&L"
    )

    # MILITARY & FOUNDATION
    p_mil = doc.add_paragraph()
    p_mil.paragraph_format.space_before = Pt(4)
    p_mil.paragraph_format.space_after = Pt(2)
    p_mil.paragraph_format.keep_with_next = True
    run_mil_h = p_mil.add_run("FOUNDATIONAL EXPERIENCE & MILITARY SERVICE")
    run_mil_h.font.name = "Arial"
    run_mil_h.font.size = Pt(8.5)
    run_mil_h.font.bold = True
    run_mil_h.font.color.rgb = navy

    add_bullet(
        "DynCorp International — Presidential Fleet (Air Force Two): Aircraft Systems Specialist, 89th Airlift Wing (1999 – 2001). Maintained Special Air Missions aircraft under strict zero-defect flight safety and operational security protocols."
    )
    add_bullet(
        "United States Air Force (USAF) — Air Force Special Operations Command (AFSOC): Aircraft Systems Specialist (1995 – 1999). Two expeditionary deployments; maintained mission-critical systems under austere operational conditions."
    )

    # 4. Education
    add_section_header("Education")
    p_edu1 = doc.add_paragraph()
    p_edu1.paragraph_format.space_before = Pt(1.5)
    p_edu1.paragraph_format.space_after = Pt(1)
    run_e1_d = p_edu1.add_run("Master of Science (M.S.) in Systems Engineering")
    run_e1_d.font.name = "Arial"
    run_e1_d.font.size = Pt(8.5)
    run_e1_d.font.bold = True
    run_e1_d.font.color.rgb = navy
    run_e1_s = p_edu1.add_run(" — Whiting School of Engineering, Johns Hopkins University (Baltimore, MD)")
    run_e1_s.font.name = "Arial"
    run_e1_s.font.size = Pt(8.5)
    run_e1_s.font.color.rgb = charcoal

    p_edu2 = doc.add_paragraph()
    p_edu2.paragraph_format.space_before = Pt(1)
    p_edu2.paragraph_format.space_after = Pt(2)
    run_e2_d = p_edu2.add_run("Bachelor of Science (B.S.) in Information Systems Science")
    run_e2_d.font.name = "Arial"
    run_e2_d.font.size = Pt(8.5)
    run_e2_d.font.bold = True
    run_e2_d.font.color.rgb = navy
    run_e2_s = p_edu2.add_run(" — University of Maryland (College Park, MD)")
    run_e2_s.font.name = "Arial"
    run_e2_s.font.size = Pt(8.5)
    run_e2_s.font.color.rgb = charcoal

    # 5. Executive Certifications & Affiliations
    add_section_header("Executive Certifications & Professional Affiliations")
    add_bullet(
        "Certified Automation Professional (CAP), International Society of Automation (ISA) • TÜV Certified Functional Safety Engineer (TÜV Rheinland / IEC 61508) • Six Sigma Master Black Belt (SSMBB - Lean & Process Quality) • Certified Digital Systems & Nuclear Cybersecurity Engineer (NANTEL / NEI) • FAA Airframe & Powerplant (A&P) License",
        "Certifications"
    )
    add_bullet(
        "Contributing Co-Author, NIST SP 800-82 Rev 3 (Guide to Operational Technology Security) • Voting Committee Member, ANSI/ISA 99 & IEC 62443 (Industrial Automation & Control Systems Security) • Voting Committee Member, ANSI/ISA 84 (Safety Instrumented Systems / SIS)",
        "Standards Bodies"
    )
    add_bullet(
        "English (Native) • Japanese (Intermediate Mid, ACTFL Scale)",
        "Languages"
    )

    output_path = r"C:\Users\mthom\.gemini\antigravity\scratch\gevernova_sr_director_ot_cybersecurity\Michael_Thompson_GE_Vernova_Sr_Director_Resume.docx"
    doc.save(output_path)
    print(f"Resume saved successfully to {output_path}")

if __name__ == "__main__":
    create_resume()
