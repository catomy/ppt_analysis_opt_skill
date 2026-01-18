# PPT Logic Analyzer

AI-powered PowerPoint logic analyzer using Pyramid Principle (金字塔原理) to analyze presentation structure and generate optimized PPTX files.

## Quick Start

```python
# 1. Extract PPT content
from utils.text_extraction import PPTXTextExtractor
extractor = PPTXTextExtractor("presentation.pptx")
data = extractor.extract_all_slides()

# 2. Claude analyzes using pre-built prompt (see SKILL.md)
# Returns JSON with suggestions

# 3. Convert suggestions to modification format
from utils.prompt_template import convert_suggestions_to_modifications
import json

suggestions = result['suggestions']
modifications = json.loads(convert_suggestions_to_modifications(suggestions))

# 4. Apply modifications
from scripts.modify_ppt import apply_modifications

prs = apply_modifications("presentation.pptx", modifications)
prs.save("presentation_optimized.pptx")
```

## Features

### 1 Analysis Dimensions

1. **逻辑结构** - Pyramid Principle (结论先行、MECE、逻辑递进）
2. **语言表达** - Complete sentences, professional tone
3. **文字排版** - Font sizes (24-30pt titles, 12-16pt body)
4. **内容组织** - 2-5 arguments per slide, parallel structure
5. **数据支撑** - Claims require data with sources
6. **视觉设计** - Color consistency, alignment, whitespace

### 2 Problem Types

| Type | Priority | Description |
|------|----------|-------------|
| `no_conclusion` | high | Title not a complete sentence |
| `illogical_order` | medium | Arguments lack logical order |
| `mece_overlap` | medium | Arguments overlap (violates MECE) |
| `unsupported_claim` | medium | Assertion lacks data/evidence |
| `poor_expression` | low | Inappropriate language |
| `font_issue` | low | Non-standard font sizes |
| `layout_problem` | low | Layout inconsistency |

## Project Structure

```
ppt-logic-analyzer/
├── README.md                 # This file
├── SKILL.md                  # Pre-built analysis prompt (full specifications)
├── scripts/
│   ├── extract_ppt_text.py   # CLI: Extract PPT text to JSON
│   ├── modify_ppt.py         # CLI: Apply JSON modifications
│   └── requirements.txt      # python-pptx>=0.6.21
├── utils/
│   ├── text_extraction.py     # PPTXTextExtractor class
│   └── prompt_template.py     # convert_suggestions_to_modifications()
└── references/
    └── pyramid_principle.md  # Pyramid Principle reference
```

## CLI Tools

```bash
# Extract PPT text
python scripts/extract_ppt_text.py presentation.pptx extracted.json
python scripts/extract_ppt_text.py presentation.pptx - > extracted.json

# Apply modifications
python scripts/modify_ppt.py input.pptx modifications.json output.pptx
```

## Installation

```bash
pip install python-pptx>=0.6.21
```

## Output Format

```json
{
  "theme": "PPT Theme",
  "total_issues": 5,
  "high_priority_issues": 3,
  "suggestions": [
    {
      "problem_type": "no_conclusion",
      "slide_number": 1,
      "location": "第1页，标题",
      "paragraph_index": 0,
      "current_content": "销售情况",
      "problem_analysis": "标题仅是主题词，缺少明确的结论性表述",
      "modification_suggestion": "第三季度销售额达到500万，同比增长25%",
      "rationale": "明确的结论能让听众快速理解核心信息",
      "priority": "high",
      "confidence": 0.95
    }
  ],
  "report": "# PPT逻辑分析报告\n..."
}
```

## Notes

- **Analysis prompt is pre-built in SKILL.md** - Claude uses it directly
- **Only python-pptx dependency** - No NLP libraries or external APIs
- **Agent-driven** - Claude performs semantic analysis, not hardcoded rules
- **Style normalization is code-driven** - 默认会统一微软雅黑与标题字号/单行等样式
- **See SKILL.md** for detailed analysis specifications and standards
