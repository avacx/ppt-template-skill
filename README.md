# PPT Template-Based Generator

A CodeBuddy Code skill that generates new PowerPoint presentations by cloning slides from existing templates, perfectly preserving all decorative elements.

## Features

- **Perfect Style Preservation**: Backgrounds, logos, shapes, animations are all preserved
- **No Warning Dialogs**: Uses python-pptx native API, no "content has issues" popups
- **Format Retention**: Replaced text keeps original font, size, color
- **Automatic Slide Detection**: Categorizes slides as cover, toc, divider, content, ending

## Installation

### Option 1: Clone to skills directory

```bash
git clone https://github.com/YOUR_USERNAME/ppt-template-skill.git ~/.codebuddy/skills/ppt-from-template
```

### Option 2: Manual installation

1. Download or clone this repository
2. Copy the entire folder to `~/.codebuddy/skills/ppt-from-template`

### Install dependencies

```bash
pip install python-pptx
```

## Usage

Once installed, simply tell CodeBuddy:

> "I have a template at `/path/to/template.pptx`, help me create a new PPT about [your topic]"

CodeBuddy will automatically:
1. Analyze the template structure
2. Create a content plan based on your requirements
3. Generate the new PPT with your content

## Manual CLI Usage

### Analyze a template

```bash
python3 scripts/ppt_cloner.py analyze "/path/to/template.pptx"
```

### Create a new PPT

```bash
python3 scripts/ppt_cloner.py create "/path/to/template.pptx" plan.json output.pptx
```

### Content plan format (plan.json)

```json
[
    {
        "template_slide": 0,
        "replacements": {
            "Original Title": "New Title"
        }
    },
    {
        "template_slide": 1,
        "replacements": {
            "Section 1": "Introduction",
            "Section 2": "Main Content"
        }
    }
]
```

## How It Works

1. **Analyze**: Reads the template PPT and identifies slide types and replaceable text
2. **Plan**: Creates a mapping of which template slides to use and what text to replace
3. **Clone**: Uses python-pptx to selectively keep slides and apply text replacements
4. **Output**: Generates a new PPTX file with all original styling intact

## File Structure

```
ppt-from-template/
├── SKILL.md           # Skill instructions for CodeBuddy
├── README.md          # This file
├── LICENSE            # MIT License
└── scripts/
    └── ppt_cloner.py  # Core Python script
```

## Requirements

- Python 3.7+
- python-pptx >= 0.6.21

## License

MIT License - see [LICENSE](LICENSE) for details.

## Contributing

Issues and pull requests are welcome!

## Author

Created for use with CodeBuddy Code.
