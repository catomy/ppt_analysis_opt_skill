#!/usr/bin/env python3
"""
Extract text from PowerPoint presentations in paragraph-level detail.

This script extracts text from PPTX files with paragraph-level structure,
enabling precise analysis and modification.
"""

import sys
import json
from utils.text_extraction import PPTXTextExtractor


def extract_presentation_text(pptx_path: str):
    """Extract paragraph-level text from a PowerPoint presentation."""
    try:
        extractor = PPTXTextExtractor(pptx_path)
        return extractor.extract_all_slides()
    except Exception as e:
        return {'error': f'Failed to open presentation: {str(e)}'}


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(json.dumps({'error': 'Usage: extract_ppt_text.py <pptx_file_path> [output_path]'}))
        sys.exit(1)

    pptx_path = sys.argv[1]
    result = extract_presentation_text(pptx_path)

    if 'error' in result:
        print(json.dumps(result))
        sys.exit(1)

    output_path = sys.argv[2] if len(sys.argv) > 2 else pptx_path.replace('.pptx', '_extracted.json')
    if output_path == '-':
        print(json.dumps(result, ensure_ascii=False, indent=2))
        sys.exit(0)

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"Extracted text saved to: {output_path}", file=sys.stderr)
