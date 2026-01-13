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

### Step 1: Analyze the Template Structure

When the user provides a template path, first analyze it to understand what slides are available and what text they contain:

```bash
python3 {{SKILL_PATH}}/scripts/ppt_cloner.py analyze "/path/to/template.pptx"
```

**Crucial**: Pay attention to the "Slide Types" (cover, toc, divider, content) and the "Text Elements" found on each slide. You will need these to map your new content.

### Step 2: Generate New Content (AI's Main Task)

Based on the user's topic, you MUST generate a complete, structured outline. Do not just replace titles; transform the entire presentation.

Example for "AI Introduction":
- **Slide 1 (Cover)**: Title: "The Future of AI", Subtitle: "Understanding Intelligence in the Digital Age"
- **Slide 2 (TOC)**: Points: "1. Definition, 2. History, 3. Main Technologies, 4. Ethics"
- **Slide 3 (Divider)**: Title: "Chapter 1: Foundations"
- **Slide 4 (Content)**: Title: "What is AI?", Body: "Artificial Intelligence is the simulation of human intelligence..."

### Step 3: Map Content to Template

Create a `plan.json` that maps your generated content to the specific strings found in Step 1.

**Strategy for "Full Replacement":**
- Find the most prominent text in a template slide (e.g., "Company Profile") and replace it with your new slide title (e.g., "What is AI?").
- Replace body text placeholders (e.g., "Enter description here") with your detailed content.
- Use the `shape:Name` syntax if you need to target a specific box precisely (found in analysis).

```json
[
    {
        "template_slide": 0,
        "replacements": {
            "Original Template Title": "The Future of AI",
            "Subtitle Placeholder": "Understanding Intelligence in the Digital Age"
        }
    }
]
```

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

## Comprehensive Transformation Tips

To ensure a "Safe" and complete replacement where no old template content remains:

1.  **Exhaustive Mapping**: Ensure every text string you see in the `analyze` output for a slide is included in your `replacements` dictionary for that slide. If you don't have new content for a specific box, replace its text with an empty string `""` or a space `" "`.
2.  **Slide Selection**: If the template has 50 slides but you only need 5, only include those 5 indices in your `plan.json`. The generator will automatically delete all other slides.
3.  **Consistency**: Use the same "Divider" slide style for all your chapter slides by referencing the same `template_slide` index multiple times in your plan.

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
