#!/usr/bin/env python3
"""
ATS-Friendly Resume Builder
Generates role-specific resumes in DOCX format from YAML data
"""

import yaml
import sys
import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_cell_border(cell, **kwargs):
    """
    Set cell borders (unused for now, but useful for tables if needed)
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'tc{}'.format(edge.capitalize())
            element = OxmlElement('w:{}'.format(tag))
            for key in ["sz", "val", "color", "space"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))
            tcPr.append(element)


def add_hyperlink(paragraph, text, url):
    """
    Add a hyperlink to a paragraph with underline formatting
    """
    # Get the part
    part = paragraph.part
    r_id = part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
    
    # Create hyperlink
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    
    # Create run
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    
    # Set underline
    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)
    
    # Set color to blue (optional, for visual hyperlink)
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0000FF')
    rPr.append(color)
    
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
    
    paragraph._p.append(hyperlink)
    return hyperlink


def filter_by_tags(items, role):
    """
    Filter items based on role tags
    """
    if role == "all":
        return items
    
    filtered = []
    for item in items:
        tags = item.get("tags", [])
        if "all" in tags or role in tags:
            filtered.append(item)
    return filtered


def apply_bold_to_text(paragraph, text):
    """
    Apply bold formatting to text within ** markers
    Example: "Increased by **27%**" -> "Increased by 27%" with 27% bolded
    """
    import re
    
    # Split text by bold markers
    parts = re.split(r'\*\*(.*?)\*\*', text)
    
    for i, part in enumerate(parts):
        if not part:
            continue
        run = paragraph.add_run(part)
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
        
        # Odd indices are the bolded parts
        if i % 2 == 1:
            run.bold = True


def build_resume(role):
    """
    Main function to build resume DOCX for a specific role
    """
    
    # Load YAML data
    with open("data/resume.yaml", "r") as f:
        data = yaml.safe_load(f)
    
    # Create document
    doc = Document()
    
    # ===========================
    # DOCUMENT MARGINS (0.5 inch)
    # ===========================
    section = doc.sections[0]
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    
    # ===========================
    # HEADER SECTION
    # ===========================
    personal = data["personal"]
    
    # Name and Title (Bold, 14pt)
    header = doc.add_paragraph()
    header.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    name_run = header.add_run(f"{personal['name']} | {personal['title']}")
    name_run.bold = True
    name_run.font.name = 'Calibri'
    name_run.font.size = Pt(14)
    header.paragraph_format.space_after = Pt(2)
    header.paragraph_format.space_before = Pt(0)
    
    # Contact Info (Regular, 10pt)
    contact = doc.add_paragraph()
    contact.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    contact_text = f"{personal['location']} | {personal['email']} | {personal['phone']} | "
    contact_run = contact.add_run(contact_text)
    contact_run.font.name = 'Calibri'
    contact_run.font.size = Pt(10)
    
    # Add hyperlinks for LinkedIn, GitHub, Website
    add_hyperlink(contact, personal['linkedin'], f"https://{personal['linkedin']}")
    contact.add_run(" | ").font.size = Pt(10)
    add_hyperlink(contact, personal['github'], f"https://{personal['github']}")
    contact.add_run(" | ").font.size = Pt(10)
    add_hyperlink(contact, personal['website'], f"https://{personal['website']}")
    
    contact.paragraph_format.space_after = Pt(8)
    contact.paragraph_format.space_before = Pt(0)
    
    # ===========================
    # SUMMARY SECTION
    # ===========================
    summary_items = filter_by_tags(data["summary"], role)
    
    if summary_items:
        # Section Header
        summary_header = doc.add_paragraph()
        summary_header_run = summary_header.add_run("SUMMARY")
        summary_header_run.bold = True
        summary_header_run.font.name = 'Calibri'
        summary_header_run.font.size = Pt(12)
        summary_header.paragraph_format.space_after = Pt(4)
        summary_header.paragraph_format.space_before = Pt(6)
        
        # Summary text as continuous paragraph
        summary_para = doc.add_paragraph()
        summary_text = " ".join([item["text"] for item in summary_items])
        summary_run = summary_para.add_run(summary_text)
        summary_run.font.name = 'Calibri'
        summary_run.font.size = Pt(11)
        summary_para.paragraph_format.space_after = Pt(8)
        summary_para.paragraph_format.line_spacing = 1.15
    
    # ===========================
    # EXPERIENCE SECTION
    # ===========================
    experience_items = []
    for exp in data["experience"]:
        filtered_bullets = filter_by_tags(exp["bullets"], role)
        if filtered_bullets:
            experience_items.append({
                "company": exp["company"],
                "role": exp["role"],
                "location": exp["location"],
                "dates": exp["dates"],
                "bullets": filtered_bullets
            })
    
    if experience_items:
        # Section Header
        exp_header = doc.add_paragraph()
        exp_header_run = exp_header.add_run("EXPERIENCE")
        exp_header_run.bold = True
        exp_header_run.font.name = 'Calibri'
        exp_header_run.font.size = Pt(12)
        exp_header.paragraph_format.space_after = Pt(4)
        exp_header.paragraph_format.space_before = Pt(6)
        
        for exp in experience_items:
            # Job Title Line: Role | Company (Location) Dates
            job_para = doc.add_paragraph()
            job_run = job_para.add_run(f"{exp['role']} | {exp['company']}")
            job_run.bold = True
            job_run.font.name = 'Calibri'
            job_run.font.size = Pt(11)
            
            location_run = job_para.add_run(f" ({exp['location']}) ")
            location_run.font.name = 'Calibri'
            location_run.font.size = Pt(11)
            
            dates_run = job_para.add_run(exp['dates'])
            dates_run.bold = True
            dates_run.font.name = 'Calibri'
            dates_run.font.size = Pt(11)
            
            job_para.paragraph_format.space_after = Pt(2)
            job_para.paragraph_format.space_before = Pt(6)
            
            # Bullets
            for bullet in exp["bullets"]:
                bullet_para = doc.add_paragraph(style='List Bullet')
                apply_bold_to_text(bullet_para, bullet["text"])
                bullet_para.paragraph_format.space_after = Pt(2)
                bullet_para.paragraph_format.line_spacing = 1.15
                bullet_para.paragraph_format.left_indent = Inches(0.25)
    
    # ===========================
    # PROJECTS SECTION
    # ===========================
    project_items = []
    for proj in data["projects"]:
        filtered_bullets = filter_by_tags(proj["bullets"], role)
        if filtered_bullets:
            project_items.append({
                "name": proj["name"],
                "tech_stack": proj["tech_stack"],
                "dates": proj["dates"],
                "bullets": filtered_bullets
            })
    
    if project_items:
        # Section Header
        proj_header = doc.add_paragraph()
        proj_header_run = proj_header.add_run("PROJECTS")
        proj_header_run.bold = True
        proj_header_run.font.name = 'Calibri'
        proj_header_run.font.size = Pt(12)
        proj_header.paragraph_format.space_after = Pt(4)
        proj_header.paragraph_format.space_before = Pt(6)
        
        for proj in project_items:
            # Project Title Line: Name | Tech Stack Dates
            proj_para = doc.add_paragraph()
            name_run = proj_para.add_run(f"{proj['name']} | ")
            name_run.bold = True
            name_run.font.name = 'Calibri'
            name_run.font.size = Pt(11)
            
            tech_run = proj_para.add_run(proj['tech_stack'])
            tech_run.font.name = 'Calibri'
            tech_run.font.size = Pt(11)
            
            dates_run = proj_para.add_run(f" {proj['dates']}")
            dates_run.bold = True
            dates_run.font.name = 'Calibri'
            dates_run.font.size = Pt(11)
            
            proj_para.paragraph_format.space_after = Pt(2)
            proj_para.paragraph_format.space_before = Pt(6)
            
            # Bullets
            for bullet in proj["bullets"]:
                bullet_para = doc.add_paragraph(style='List Bullet')
                apply_bold_to_text(bullet_para, bullet["text"])
                bullet_para.paragraph_format.space_after = Pt(2)
                bullet_para.paragraph_format.line_spacing = 1.15
                bullet_para.paragraph_format.left_indent = Inches(0.25)
    
    # ===========================
    # EDUCATION SECTION
    # ===========================
    education_items = filter_by_tags(data["education"], role)
    
    if education_items:
        # Section Header
        edu_header = doc.add_paragraph()
        edu_header_run = edu_header.add_run("EDUCATION")
        edu_header_run.bold = True
        edu_header_run.font.name = 'Calibri'
        edu_header_run.font.size = Pt(12)
        edu_header.paragraph_format.space_after = Pt(4)
        edu_header.paragraph_format.space_before = Pt(6)
        
        for edu in education_items:
            # Degree Line (Bold)
            degree_para = doc.add_paragraph()
            degree_run = degree_para.add_run(f"{edu['degree']} {edu['dates']}")
            degree_run.bold = True
            degree_run.font.name = 'Calibri'
            degree_run.font.size = Pt(11)
            degree_para.paragraph_format.space_after = Pt(2)
            degree_para.paragraph_format.space_before = Pt(6)
            
            # Institution Line (Regular)
            inst_para = doc.add_paragraph()
            inst_run = inst_para.add_run(f"{edu['institution']} ({edu['location']})")
            inst_run.font.name = 'Calibri'
            inst_run.font.size = Pt(11)
            inst_para.paragraph_format.space_after = Pt(4)
    
    # ===========================
    # SKILLS SECTION
    # ===========================
    skills_data = data["skills"]
    skill_items = []
    
    for key, skill_cat in skills_data.items():
        tags = skill_cat.get("tags", [])
        if "all" in tags or role in tags:
            skill_items.append(skill_cat)
    
    if skill_items:
        # Section Header
        skills_header = doc.add_paragraph()
        skills_header_run = skills_header.add_run("SKILLS")
        skills_header_run.bold = True
        skills_header_run.font.name = 'Calibri'
        skills_header_run.font.size = Pt(12)
        skills_header.paragraph_format.space_after = Pt(4)
        skills_header.paragraph_format.space_before = Pt(6)
        
        for skill_cat in skill_items:
            skill_para = doc.add_paragraph()
            label_run = skill_para.add_run(f"{skill_cat['label']}: ")
            label_run.bold = True
            label_run.font.name = 'Calibri'
            label_run.font.size = Pt(11)
            
            items_run = skill_para.add_run(", ".join(skill_cat['items']))
            items_run.font.name = 'Calibri'
            items_run.font.size = Pt(11)
            
            skill_para.paragraph_format.space_after = Pt(2)
            skill_para.paragraph_format.line_spacing = 1.15
    
    # ===========================
    # SAVE DOCUMENT
    # ===========================
    os.makedirs("outputs", exist_ok=True)
    output_path = f"outputs/resume_{role}.docx"
    doc.save(output_path)
    
    print(f"✅ Generated resume for role: {role}")
    print(f"📄 Output: {output_path}")
    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python build_docx.py <role>")
        print("Example: python build_docx.py python")
        print("Available roles: python, java, fullstack, backend, frontend, cloud, devops, mobile, all")
        sys.exit(1)
    
    role = sys.argv[1].lower()
    build_resume(role)
