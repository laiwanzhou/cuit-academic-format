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


def test_reference_heuristic_checks_in_references_only():
    mod = load_module()
    doc = mod.Document()
    doc.add_paragraph("references")
    doc.add_paragraph("[1] 作者. 题名. https://example.com")
    doc.add_paragraph("[3] 作者. 题名[J]. 期刊名")
    mod.analyze_section_sequence = lambda texts, has_toc_field=False: {"regions": ["references"] * len(texts)}
    issues = mod.collect_reference_issues(doc)
    keys = {x.rule_key for x in issues}
    assert "reference_online_access" in keys
    assert "reference_sequence" in keys


def test_reference_heuristic_detection_cases():
    mod = load_module()
    doc = mod.Document()
    doc.add_paragraph("references")
    doc.add_paragraph("[1] 作者. 题名. 期刊名, 2020, 12(3): 15-20.")
    doc.add_paragraph("[3] 作者. 题名[J]. 期刊名")
    doc.add_paragraph("[4] 作者. 题名[EB/OL]. (2020-01-01). https://example.com.")
    doc.add_paragraph("[5] 作者. 题名[D]. 成都: 成都信息工程大学, 2023.")
    mod.analyze_section_sequence = lambda texts, has_toc_field=False: {"regions": ["references"] * len(texts)}
    issues = mod.collect_reference_issues(doc)
    keys = [x.rule_key for x in issues]
    assert "reference_type_marker" in keys
    assert "reference_sequence" in keys
    assert "reference_online_access" in keys
    assert "reference_entry_format" in keys


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
