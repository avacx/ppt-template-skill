---
name: ppt-from-template
description: "Generate new PowerPoint presentations based on existing templates. When users provide a template PPT file path and want to create new presentations that preserve all decorative elements (backgrounds, logos, headers/footers) while replacing text content. Triggers: (1) User provides a .pptx template path, (2) User asks to create PPT based on existing style, (3) User wants to replicate PPT design with new content."
---

# PPT Template-Based Generator

Create new PowerPoint presentations by cloning slides from an existing template and replacing text content. This approach **perfectly preserves** all decorative elements including backgrounds, logos, shapes, and formatting.

## When to Use This Skill

- User provides a local `.pptx` file path as a template
- User wants to create a new PPT with the same style/design as an existing one
- User asks to "use this PPT as a template" or "replicate this PPT style"

## Workflow

### Step 1: Analyze the Template

When the user provides a template path, first analyze its structure:

```bash
python3 {{SKILL_PATH}}/scripts/ppt_cloner.py analyze "/path/to/template.pptx"
```

This outputs:
- Total slide count
- Slide types (cover, toc, divider, content, ending)
- Text elements that can be replaced on each slide
- Preview of existing content

### Step 2: Plan the Content

Based on the analysis and user requirements, create a content plan JSON file:

```json
[
    {
        "template_slide": 0,
        "replacements": {
            "Original Title": "New Title",
            "Original Subtitle": "New Subtitle"
        }
    },
    {
        "template_slide": 1,
        "replacements": {
            "Chapter 1": "Introduction",
            "Chapter 2": "Main Content"
        }
    }
]
```

**Fields:**
- `template_slide`: Which slide from the template to use (0-indexed)
- `replacements`: Text replacement rules (exact match required)

### Step 3: Generate the PPT

```bash
python3 {{SKILL_PATH}}/scripts/ppt_cloner.py create "/path/to/template.pptx" content_plan.json "/path/to/output.pptx"
```

## Complete Example

**User Request:**
> I have a template at `/Users/xxx/company_template.pptx`. Create a quarterly report with:
> - Cover: Q4 2024 Report
> - Sections: Sales Performance, Product Updates, Next Quarter Plans
> - Ending: Thank You

**AI Execution:**

1. **Analyze template:**
```bash
python3 {{SKILL_PATH}}/scripts/ppt_cloner.py analyze "/Users/xxx/company_template.pptx"
```

2. **Review output to understand:**
   - Which slides are cover, divider, content types
   - What text can be replaced

3. **Create `plan.json`:**
```json
[
    {
        "template_slide": 0,
        "replacements": {
            "Company Template": "Q4 2024 Report",
            "Subtitle Here": "Quarterly Business Review"
        }
    },
    {
        "template_slide": 1,
        "replacements": {
            "Section 1": "Sales Performance",
            "Section 2": "Product Updates",
            "Section 3": "Next Quarter Plans"
        }
    },
    {
        "template_slide": 2,
        "replacements": {
            "01": "01",
            "Section Title": "Sales Performance"
        }
    }
]
```

4. **Generate PPT:**
```bash
python3 {{SKILL_PATH}}/scripts/ppt_cloner.py create "/Users/xxx/company_template.pptx" plan.json "/Users/xxx/Q4_Report.pptx"
```

5. **Inform user:** "Generated `/Users/xxx/Q4_Report.pptx`"

## Key Features

| Feature | Description |
|---------|-------------|
| **Perfect Style Preservation** | All backgrounds, logos, shapes, animations are preserved |
| **No Warning Dialogs** | Uses python-pptx native API, no "content has issues" popups |
| **Format Retention** | Replaced text keeps original font, size, color |
| **Flexible Selection** | Choose any slides from template in any order |

## Slide Type Detection

The analyzer automatically categorizes slides:

| Type | Detection Criteria |
|------|---------------------|
| `cover` | First slide, or layout name contains "title/cover" |
| `toc` | Contains "目录/contents/agenda" text |
| `divider` | Few text elements with standalone numbers (01, 02...) |
| `content` | Default type for regular content slides |
| `ending` | Contains "谢谢/thank/感谢" text |

## Text Replacement Rules

1. **Exact Match Required**: The `old_text` must exactly match content in the slide
2. **Partial Text OK**: If a text box contains "Hello World", you can replace "Hello" only
3. **Multiple Replacements**: All matching occurrences in the slide will be replaced
4. **Preserve Formatting**: Font, size, color, bold/italic are preserved

## Dependencies

Requires `python-pptx`:

```bash
pip install python-pptx
```

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Text not replaced | Check exact spelling including spaces |
| Wrong slide used | Verify `template_slide` index (0-based) |
| Missing elements | Ensure template file path is correct |

## Script Location

The core script is at: `{{SKILL_PATH}}/scripts/ppt_cloner.py`
