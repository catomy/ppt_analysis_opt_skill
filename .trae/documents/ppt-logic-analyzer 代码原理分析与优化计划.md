## 代码原理（当前实现的数据流）
- 提取：通过 [text_extraction.py](file:///c:/Users/piaom/.claude/skills/ppt-logic-analyzer/utils/text_extraction.py) 的 `PPTXTextExtractor` 读取 pptx，按“上→下、左→右”对形状排序，尝试识别标题形状，然后以“形状→段落→runs(字体信息)”的层级抽取结构化数据。
- 分析：分析逻辑不在代码中实现，而是写在 [SKILL.md](file:///c:/Users/piaom/.claude/skills/ppt-logic-analyzer/SKILL.md) 里的长 Prompt，由 Claude 依据抽取结果输出 JSON suggestions（含 `slide_number/location/paragraph_index/modification_suggestion` 等）。
- 转换：用 [prompt_template.py](file:///c:/Users/piaom/.claude/skills/ppt-logic-analyzer/utils/prompt_template.py) 的 `convert_suggestions_to_modifications()` 把 suggestions 按页分组，生成 `modify_ppt.py` 可消费的 modifications（`replace_by_index`）。
- 应用：用 [modify_ppt.py](file:///c:/Users/piaom/.claude/skills/ppt-logic-analyzer/scripts/modify_ppt.py) 按 slide_index 找到页，再按“段落计数”定位到目标段落，把第一段 run 的文本替换成新文本，并删掉同段落其余 runs。

## 关键问题（会导致错误/不一致/不可用）
1) `to_json()` 直接不可用（JSON 序列化必崩）
- [text_extraction.py:L21-L25](file:///c:/Users/piaom/.claude/skills/ppt-logic-analyzer/utils/text_extraction.py#L21-L25) 返回结构里包含 `file: self.prs`（Presentation 对象不可 JSON 序列化）。
- [text_extraction.py:L210-L213](file:///c:/Users/piaom/.claude/skills/ppt-logic-analyzer/utils/text_extraction.py#L210-L213) `to_json()` 会对整个 dict `json.dumps`，因此会抛 TypeError。

2) “抽取格式 / paragraph_index 语义 / 修改器定位方式”存在内在矛盾
- `utils/PPTXTextExtractor` 输出是 `content_shapes -> paragraphs` 的嵌套结构（按形状分组）。见 [text_extraction.py:L37-L58](file:///c:/Users/piaom/.claude/skills/ppt-logic-analyzer/utils/text_extraction.py#L37-L58)。
- 但 Prompt 里要求 `paragraph_index` 对应“content 数组索引”（扁平数组），见 [SKILL.md:L247-L251](file:///c:/Users/piaom/.claude/skills/ppt-logic-analyzer/SKILL.md#L247-L251)。
- `modify_ppt.py` 的替换定位也是“跳过标题 shape 后，跨 shape 累计段落数”的扁平计数法，见 [modify_ppt.py:L36-L60](file:///c:/Users/piaom/.claude/skills/ppt-logic-analyzer/scripts/modify_ppt.py#L36-L60)。
- 结果：如果分析用的是 utils 抽取结果，LLM 产出的 `paragraph_index` 很可能无法稳定映射到 `modify_ppt.py` 的扁平计数，导致改错段落。

3) 标题识别在抽取与修改两端不一致，容易“改错标题/漏改标题”
- 抽取端标题识别比较复杂（顶部+字体阈值、占位符类型、兜底第一个文本形状），见 [text_extraction.py:L129-L163](file:///c:/Users/piaom/.claude/skills/ppt-logic-analyzer/utils/text_extraction.py#L129-L163)。
- 修改端标题识别非常简化：默认“排序后第 1 个文本 shape”就是标题，见 [modify_ppt.py:L27-L35](file:///c:/Users/piaom/.claude/skills/ppt-logic-analyzer/scripts/modify_ppt.py#L27-L35)。
- 两端标准不一致时，`target_type=title` 可能会改到页眉/小标题/正文。

4) 转换器用 `location` 文本包含“标题”来判定 target，鲁棒性不足
- [prompt_template.py:L45-L51](file:///c:/Users/piaom/.claude/skills/ppt-logic-analyzer/utils/prompt_template.py#L45-L51) 仅用 `'标题' in location` 来判断是否标题；如果输出语言变化/location 格式变化/标题写成“Title”等，会误分类。

5) 修改器缺少“安全校验”，一旦索引错就会静默改错
- modifications 里带了 `old_text`（用于验证），但 `replace_by_index` 分支完全没用它做校验，见 [modify_ppt.py:L25-L60](file:///c:/Users/piaom/.claude/skills/ppt-logic-analyzer/scripts/modify_ppt.py#L25-L60)。

## 可改进点（质量、可维护性、能力边界）
- 统一“唯一的数据结构契约”：抽取 JSON 与 prompt 说明、修改器的段落定位必须同一套语义。
- 引入“稳定定位符”替代全局扁平 paragraph_index：例如 `(shape_id/shape_idx, para_idx)`，并在抽取结果中明确给出；修改时优先用稳定 id，索引仅作兼容。
- 修改前做校验：如果目标段落当前文本与 `old_text` 不一致，采取降级策略（在 slide 内查找最接近匹配文本、或拒绝修改并输出错误报告）。
- 处理表格/组合形状：抽取端已对表格做了 `table_data`，但修改器不支持按 `row/col` 替换表格单元格；组合形状抽取目前仅提取文本但没有结构化定位。
- 标题/页眉页脚判定改为占位符类型优先（DATE/FOOTER/SLIDE_NUMBER/TITLE/CENTER_TITLE 等），减少“位置阈值”误判。

## 我将实施的优化（获得你确认后才会改代码）
1) 定义并落地一个“Canonical Extracted JSON Schema”
- slide：`slide_number/slide_id/title/content(flat)/content_shapes(optional)/notes`
- content(flat) 每条包含：`global_index`（兼容旧）、`shape_index`、`shape_id(若可得)`、`paragraph_index_in_shape`、`text`、`level`、`runs`。

2) 修复 `PPTXTextExtractor.to_json()` 的序列化问题
- 将 `file` 字段改为 `pptx_path` 或移除，确保任何导出都可直接 `json.dumps`。

3) 去重并统一抽取入口
- 让 [scripts/extract_ppt_text.py](file:///c:/Users/piaom/.claude/skills/ppt-logic-analyzer/scripts/extract_ppt_text.py) 复用 `utils/PPTXTextExtractor`（或反过来保留脚本逻辑、废弃 utils 版本），避免两套格式长期漂移。

4) 改造修改器以支持“稳定定位 + 校验”
- 新增 change 类型：`replace_by_shape_paragraph`（使用 shape_index + paragraph_index_in_shape）。
- `replace_by_index` 保留做兼容，但会校验 `old_text`；不匹配则拒绝或 fallback 搜索。
- 标题修改优先使用占位符标题（若存在），否则用与抽取端一致的规则。

5) 对 Prompt/说明文档做一致性修订
- 更新 [SKILL.md](file:///c:/Users/piaom/.claude/skills/ppt-logic-analyzer/SKILL.md) 中关于 `paragraph_index` 的定义，使其与抽取 JSON 完全一致（推荐输出 shape+para 双索引），并更新示例。
- 更新 [README.md](file:///c:/Users/piaom/.claude/skills/ppt-logic-analyzer/README.md) 的示例，避免展示不可序列化对象。

6) 增加最小化自动化验证（不依赖外部 PPT 文件）
- 用 python-pptx 在测试里生成一个临时 pptx（含标题、正文、表格），跑“抽取→构造 modifications→应用→再抽取比对”，验证不会改错段落。

确认后我会按以上 6 步修改代码与文档，并提供一份兼容旧格式的迁移说明。