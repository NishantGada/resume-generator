### **{{ name }} | {{ title }}**
{{ email }} | {{ linkedin }} | {{ github }}

---

### Summary
{% for s in summary %}
- {{ s }}
{% endfor %}

---

### Experience
{% for exp in experience %}
### {{ exp.role }} – {{ exp.company }}
*{{ exp.dates }}*

{% for b in exp.bullets %}
- {{ b }}
{% endfor %}

{% endfor %}

---

### Skills
{% for category, items in skills.items() %}
**{{ category | capitalize }}:** {{ ", ".join(items) }}

{% endfor %}
