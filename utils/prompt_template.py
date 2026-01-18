"""
PPT分析辅助工具

为Claude/OpenCode agent提供的辅助函数
"""

from pathlib import Path
from typing import Dict, Any, List
import json


def convert_suggestions_to_modifications(suggestions: List[Dict[str, Any]]) -> str:
    """
    将suggestions转换为modify_ppt.py可识别的JSON格式（段落级修改）

    Args:
        suggestions: 分析返回的suggestions列表
                     每个suggestion包含:
                        - problem_type: 问题类型
                        - slide_number: 页码（1-based）
                        - location: 位置（如"第3页，标题" 或 "第5页，第8段"）
                        - paragraph_index: 段落索引（标题为0，内容从0开始）
                        - current_content: 当前内容（用于验证）
                        - modification_suggestion: 修改建议
                        - 其他字段...

    Returns:
        JSON字符串，可直接用于modify_ppt.py

    Note:
        - slide_index 使用 1-based 编号（与 slide_number 一致）
        - 使用 paragraph_index 进行精确的段落级修改
        - 按页分组，每页可以包含标题和内容的多个修改
    """
    # 按页和目标类型分组
    modifications_by_slide = {}

    for sugg in suggestions:
        slide_num = sugg.get('slide_number', 1)

        # 初始化该页的修改
        if slide_num not in modifications_by_slide:
            modifications_by_slide[slide_num] = {'title': [], 'content': []}

        # 判断是标题还是内容
        location = str(sugg.get('location', ''))
        is_title = sugg.get('target_type') == 'title' or ('标题' in location)
        target_key = 'title' if is_title else 'content'

        # 获取段落索引（标题固定为0）
        para_index = 0 if is_title else sugg.get('paragraph_index', 0)

        new_text = sugg.get('modification_suggestion', '')
        old_text = sugg.get('current_content', '')

        can_use_stable_locator = (
            not is_title
            and isinstance(sugg.get('shape_index', None), int)
            and (
                isinstance(sugg.get('paragraph_index_in_shape', None), int)
                or isinstance(sugg.get('nonempty_index_in_shape', None), int)
            )
        )

        if can_use_stable_locator:
            change = {
                "type": "replace_by_shape_paragraph",
                "shape_index": sugg.get('shape_index'),
                "paragraph_index_in_shape": sugg.get('paragraph_index_in_shape'),
                "nonempty_index_in_shape": sugg.get('nonempty_index_in_shape'),
                "paragraph_index": para_index,
                "new_text": new_text,
                "old_text": old_text
            }
        else:
            change = {
                "type": "replace_by_index",
                "paragraph_index": para_index,
                "new_text": new_text,
                "old_text": old_text
            }

        modifications_by_slide[slide_num][target_key].append(change)

    # 转换为最终格式
    modifications = []

    for slide_num in sorted(modifications_by_slide.keys()):
        slide_mods = modifications_by_slide[slide_num]

        # 添加标题修改
        if slide_mods['title']:
            modifications.append({
                "slide_index": slide_num,
                "target_type": "title",
                "changes": slide_mods['title']
            })

        # 添加内容修改
        if slide_mods['content']:
            modifications.append({
                "slide_index": slide_num,
                "target_type": "content",
                "changes": slide_mods['content']
            })

    return json.dumps(modifications, ensure_ascii=False, indent=2)


def prepare_suggestions_for_refine(suggestions: List[Dict[str, Any]]) -> str:
    refined_keys = [
        'problem_type',
        'slide_number',
        'location',
        'paragraph_index',
        'shape_index',
        'paragraph_index_in_shape',
        'nonempty_index_in_shape',
        'current_content',
        'modification_suggestion',
        'priority',
        'confidence',
    ]

    normalized = []
    for suggestion in suggestions:
        normalized.append({k: suggestion.get(k) for k in refined_keys if k in suggestion})

    return json.dumps(normalized, ensure_ascii=False, indent=2)


def merge_refined_suggestions(
    original_suggestions: List[Dict[str, Any]],
    refined_suggestions: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    refined_by_key = {}
    for suggestion in refined_suggestions:
        key = (
            suggestion.get('slide_number'),
            suggestion.get('location'),
            suggestion.get('paragraph_index'),
            suggestion.get('current_content'),
        )
        refined_by_key[key] = suggestion.get('modification_suggestion')

    merged = []
    for original in original_suggestions:
        key = (
            original.get('slide_number'),
            original.get('location'),
            original.get('paragraph_index'),
            original.get('current_content'),
        )
        new_modification_suggestion = refined_by_key.get(key, None)
        if new_modification_suggestion is None:
            merged.append(original)
            continue

        updated = dict(original)
        updated['modification_suggestion'] = new_modification_suggestion
        merged.append(updated)

    return merged


def get_pyramid_principle_reference() -> str:
    """
    获取金字塔原理参考文档

    Returns:
        金字塔原理详细说明（来自references/pyramid_principle.md）
    """
    pyramid_path = Path(__file__).parent.parent / "references" / "pyramid_principle.md"

    if pyramid_path.exists():
        with open(pyramid_path, 'r', encoding='utf-8') as f:
            return f.read()
    else:
        return "参考文档未找到，请确保references/pyramid_principle.md文件存在"
