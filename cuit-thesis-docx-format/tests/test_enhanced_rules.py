import importlib.util
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def load_module():
    module_path = ROOT / "scripts" / "cuit_thesis_docx_format.py"
    spec = importlib.util.spec_from_file_location("cuit_mod", module_path)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


def test_body_alignment_and_caption_line_spacing_rules():
    mod = load_module()
    assert mod.RULES["body"].alignment == mod.ALIGN_JUSTIFIED
    assert mod.RULES["figure_caption"].space_after_lines == 1
    assert mod.RULES["table_caption"].space_before_lines == 1


def test_body_cjk_spacing_only_for_body():
    mod = load_module()
    doc = mod.Document()
    doc.add_paragraph("摘 要")
    doc.add_paragraph("在 Web 层")
    doc.add_paragraph("第一章 绪论")
    doc.add_paragraph("在 Web 实现层面， Flask 由于核心简洁")
    issues = mod.collect_issues(doc)
    spacing = [x for x in issues if x.rule_key == "body_cjk_spacing"]
    assert len(spacing) >= 1


def test_section_boundary_extra_author_bio_warning_not_module():
    mod = load_module()
    texts = [
        "致谢",
        "致谢",
        "致谢",
        "第一章 绪论",
        "无关段落",
        "致谢",
        "作者简历及攻读学位期间发表的学术论文与研究成果",
    ]
    regions = mod.analyze_section_sequence(texts)["regions"]
    errors, warnings = mod.compute_section_boundary_findings(texts, regions)
    assert any("疑似额外组成部分" in x for x in warnings)
    assert all("author_bio" not in x for x in errors)


def _collect_reference_keys(mod, entries):
    doc = mod.Document()
    doc.add_paragraph("参考文献")
    for line in entries:
        doc.add_paragraph(line)
    mod.analyze_section_sequence = lambda texts, has_toc_field=False: {"regions": ["references"] * len(texts)}
    issues = mod.collect_reference_issues(doc)
    return issues, {x.rule_key for x in issues}


def test_reference_272_valid_journal_match():
    mod = load_module()
    issues, keys = _collect_reference_keys(mod, ["[1] 作者. 题名[J]. 期刊名, 2020, 12(3): 15-20."])
    assert "reference_272_format_no_match" not in keys
    assert "reference_272_type_format_mismatch" not in keys


def test_reference_272_invalid_journal():
    mod = load_module()
    issues, keys = _collect_reference_keys(mod, ["[1] 作者. 题名[J]. 期刊名."])
    assert "reference_272_type_format_mismatch" in keys or "reference_entry_format" in keys


def test_reference_272_valid_dissertation():
    mod = load_module()
    _issues, keys = _collect_reference_keys(mod, ["[1] 作者. 题名[D]. 成都: 成都信息工程大学, 2023."])
    assert "reference_272_format_no_match" not in keys


def test_reference_272_invalid_dissertation():
    mod = load_module()
    _issues, keys = _collect_reference_keys(mod, ["[1] 作者. 题名[D]."])
    assert "reference_272_type_format_mismatch" in keys


def test_reference_272_valid_ebol():
    mod = load_module()
    _issues, keys = _collect_reference_keys(mod, ["[1] 作者. 题名[EB/OL]. (2020-01-01)[2024-01-01]. https://example.com."])
    assert "reference_272_format_no_match" not in keys


def test_reference_272_invalid_ebol_missing_citation_date():
    mod = load_module()
    _issues, keys = _collect_reference_keys(mod, ["[1] 作者. 题名[EB/OL]. https://example.com."])
    assert "reference_273_online_access_missing" in keys or "reference_272_type_format_mismatch" in keys


def test_reference_272_no_match():
    mod = load_module()
    _issues, keys = _collect_reference_keys(mod, ["[1] 这是一条随便写的参考文献"])
    assert "reference_272_format_no_match" in keys


def test_reference_missing_type_marker():
    mod = load_module()
    _issues, keys = _collect_reference_keys(mod, ["[1] 作者. 题名. 期刊名, 2020, 12(3): 15-20."])
    assert "reference_type_marker" in keys


def test_reference_sequence_not_continuous():
    mod = load_module()
    _issues, keys = _collect_reference_keys(
        mod,
        [
            "[1] 作者. 题名[J]. 期刊名, 2020, 12(3): 15-20.",
            "[3] 作者. 题名[J]. 期刊名, 2021, 13(4): 21-30.",
        ],
    )
    assert "reference_sequence" in keys


def test_reference_272_valid_m_s_p_n():
    mod = load_module()
    _issues, keys = _collect_reference_keys(
        mod,
        [
            "[1] 作者. 题名[M]. 北京: 出版社, 2020: 10-20.",
            "[2] GB/T 7714-2015, 信息与文献 参考文献著录规则[S]. 北京: 中国标准出版社, 2015.",
            "[3] 张三. 一种数据处理方法: CN101234567A[P]. 2020-01-01.",
            "[4] 作者. 题名[N]. 人民日报, 2020-01-01(01).",
        ],
    )
    assert "reference_272_format_no_match" not in keys


def test_reference_273_online_missing_for_journal_with_url():
    mod = load_module()
    _issues, keys = _collect_reference_keys(mod, ["[1] 作者. 题名[J]. 期刊名, 2020, 12(3): 15-20. https://example.com."])
    assert "reference_273_online_access_missing" in keys


def test_reference_273_author_et_al():
    mod = load_module()
    _issues, keys = _collect_reference_keys(mod, ["[1] 张三, 李四, 王五, 赵六. 题名[J]. 期刊名, 2020, 12(3): 15-20."])
    assert "reference_273_author_et_al" in keys


def test_reference_checks_scope_only_references():
    mod = load_module()
    doc = mod.Document()
    doc.add_paragraph("第一章 绪论")
    doc.add_paragraph("[1] 作者. 题名[J]. 期刊名, 2020, 12(3): 15-20.")
    issues = mod.collect_reference_issues(doc)
    assert issues == []


def test_empty_paragraph_cleanup_targets():
    mod = load_module()
    doc = mod.Document()
    doc.add_paragraph("摘 要")
    doc.add_paragraph("   ")
    doc.add_paragraph("ABSTRACT")
    doc.add_paragraph("　")
    doc.add_paragraph("第一章 绪论")
    doc.add_paragraph("	")
    doc.add_paragraph("致谢")
    doc.add_paragraph("  ")
    issues, removable = mod.collect_empty_paragraph_issues(doc)
    assert len(issues) >= 1
    removed = mod.remove_target_empty_paragraphs(doc, removable)
    assert removed >= 1


def test_image_size_advisory_and_height_limit():
    mod = load_module()

    class FakeDim:
        def __init__(self, cm):
            self.cm = cm

    class FakeShape:
        def __init__(self, w, h):
            self.width = FakeDim(w)
            self.height = FakeDim(h)

    class FakeDoc:
        def __init__(self, shapes):
            self.inline_shapes = shapes

    doc = FakeDoc([FakeShape(8, 6), FakeShape(10, 12), FakeShape(9, 20)])
    issues = mod.collect_image_size_issues(doc)
    keys = [x.rule_key for x in issues]
    assert "image_size_advisory" in keys
    assert "image_height_limit" in keys


def test_table_caption_and_three_line_checks():
    mod = load_module()
    doc = mod.Document()
    doc.add_paragraph("无关段落")
    table = doc.add_table(rows=2, cols=2)
    table.style = "Table Grid"
    doc.add_paragraph("表2-1 示例表题")
    issues = mod.collect_table_issues(doc)
    keys = {x.rule_key for x in issues}
    assert "table_caption_missing" in keys
    assert "table_caption_position" in keys
    assert "table_three_line_style" in keys


def test_english_abstract_title_normalization_and_issue():
    mod = load_module()
    doc = mod.Document()
    doc.add_paragraph("Abstract")
    doc.add_paragraph("This is abstract body.")
    mod.analyze_section_sequence = lambda texts, has_toc_field=False: {"regions": ["abstract_en"] * len(texts)}
    mod.normalize_english_abstract_title(doc)
    assert doc.paragraphs[0].text == "ABSTRACT"
    issues = mod.collect_abstract_title_issues(doc)
    keys = {x.rule_key for x in issues}
    assert "abstract_title_en_text" not in keys


def test_english_abstract_title_colon_normalization():
    mod = load_module()
    doc = mod.Document()
    doc.add_paragraph("ABSTRACT:")
    doc.add_paragraph("Body line.")
    mod.analyze_section_sequence = lambda texts, has_toc_field=False: {"regions": ["abstract_en"] * len(texts)}
    mod.normalize_english_abstract_title(doc)
    assert doc.paragraphs[0].text == "ABSTRACT"


def test_english_abstract_title_mixed_with_body_report_only():
    mod = load_module()
    doc = mod.Document()
    doc.add_paragraph("Abstract: With the outbreak of COVID-19...")
    mod.analyze_section_sequence = lambda texts, has_toc_field=False: {"regions": ["abstract_en"]}
    before = doc.paragraphs[0].text
    mod.normalize_english_abstract_title(doc)
    assert doc.paragraphs[0].text == before
    issues = mod.collect_abstract_title_issues(doc)
    assert any(issue.rule_key == "abstract_title_en_text" for issue in issues)


def test_body_abstract_word_not_reported_as_title():
    mod = load_module()
    doc = mod.Document()
    doc.add_paragraph("This section discusses Abstract Factory pattern.")
    mod.analyze_section_sequence = lambda texts, has_toc_field=False: {"regions": ["body"]}
    issues = mod.collect_abstract_title_issues(doc)
    assert issues == []


def test_table_caption_number_formats_and_spacing():
    mod = load_module()
    doc = mod.Document()
    doc.add_paragraph("表2-1 示例表题")
    t1 = doc.add_table(rows=2, cols=2)
    t1.style = "Normal Table"
    doc.add_paragraph("表1 示例表题")
    t2 = doc.add_table(rows=2, cols=2)
    t2.style = "Normal Table"
    doc.add_paragraph("表2.1 示例表题")
    t3 = doc.add_table(rows=2, cols=2)
    t3.style = "Normal Table"
    issues = mod.collect_table_issues(doc)
    keys = [x.rule_key for x in issues]
    assert "table_caption_missing" not in keys
    assert "table_caption_number_format" not in keys


def test_table_caption_missing_space_detected():
    mod = load_module()
    doc = mod.Document()
    doc.add_paragraph("表2-1示例表题")
    table = doc.add_table(rows=2, cols=2)
    table.style = "Normal Table"
    issues = mod.collect_table_issues(doc)
    assert any(issue.rule_key == "table_caption_number_format" for issue in issues)


def test_continued_table_prompts():
    mod = load_module()
    doc = mod.Document()
    doc.add_paragraph("续表（续）")
    table = doc.add_table(rows=1, cols=2)
    table.style = "Normal Table"
    issues = mod.collect_table_issues(doc)
    assert any(issue.rule_key == "table_continued_header" for issue in issues)
