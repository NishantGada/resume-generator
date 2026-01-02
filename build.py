import yaml
import sys
from jinja2 import Template

ROLE = sys.argv[1].lower()

with open("data/resume.yaml") as f:
    data = yaml.safe_load(f)

def filter_items(items):
    return [item["text"] for item in items if ROLE in item["tags"]]

# Filter summary
summary = filter_items(data["summary"])

# Filter experience
experience = []
for exp in data["experience"]:
    bullets = filter_items(exp["bullets"])
    if bullets:
        experience.append({
            "company": exp["company"],
            "role": exp["role"],
            "dates": exp["dates"],
            "bullets": bullets
        })

# Filter skills
skills = {k: v for k, v in data["skills"].items() if k == ROLE or ROLE == "fullstack"}

# Render Markdown
with open("templates/resume.md") as f:
    tmpl = Template(f.read())

rendered_md = tmpl.render(
    name=data["name"],
    title=data["title"],
    email=data["email"],
    linkedin=data["linkedin"],
    github=data["github"],
    summary=summary,
    experience=experience,
    skills=skills
)

out_path = f"outputs/resume_{ROLE}.md"
with open(out_path, "w") as f:
    f.write(rendered_md.strip() + "\n")

print(f"✔ Generated Markdown for role: {ROLE} → {out_path}")
