import importlib.util
from pathlib import Path
import sys

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
    doc.add_paragraph("\u76ee \u5f55")
    doc.add_paragraph("\u5728 Web \u5c42")
    doc.add_paragraph("\u7b2c\u4e00\u7ae0 \u7eea\u8bba")
    doc.add_paragraph("\u5728 Web \u5b9e\u73b0\u5c42\u9762\uff0c Flask \u7531\u4e8e\u6838\u5fc3\u7b80\u6d01")
    issues = mod.collect_issues(doc)
    spacing = [x for x in issues if x.rule_key == "body_cjk_spacing"]
    assert len(spacing) >= 1


def test_section_boundary_extra_author_bio_warning_not_module():
    mod = load_module()
    texts = [
        "\u5c01\u9762",
        "\u6458\u8981",
        "\u76ee\u5f55",
        "\u7b2c\u4e00\u7ae0 \u7eea\u8bba",
        "\u53c2\u8003\u6587\u732e",
        "\u81f4\u8c22",
        "\u4f5c\u8005\u7b80\u5386\u53ca\u653b\u8bfb\u5b66\u4f4d\u671f\u95f4\u53d1\u8868\u7684\u5b66\u672f\u8bba\u6587\u4e0e\u7814\u7a76\u6210\u679c",
    ]
    regions = mod.analyze_section_sequence(texts)["regions"]
    errors, warnings = mod.compute_section_boundary_findings(texts, regions)
    assert any("疑似额外组成部分" in x for x in warnings)
    assert all("author_bio" not in x for x in errors)


def test_reference_heuristic_checks_in_references_only():
    mod = load_module()
    doc = mod.Document()
    doc.add_paragraph("references header")
    doc.add_paragraph("[1] ??. ??. https://example.com")
    doc.add_paragraph("[3] ??. Another Title[J]. Journal")
    mod.analyze_section_sequence = lambda texts, has_toc_field=False: {"regions": ["references"] * len(texts)}
    issues = mod.collect_reference_issues(doc)
    keys = {x.rule_key for x in issues}
    assert "reference_online_access" in keys
    assert "reference_sequence" in keys


def test_empty_paragraph_cleanup_targets():
    mod = load_module()
    doc = mod.Document()
    doc.add_paragraph("? ?")
    doc.add_paragraph("   ")
    doc.add_paragraph("ABSTRACT")
    doc.add_paragraph("　")
    doc.add_paragraph("??? ??")
    doc.add_paragraph("	")
    doc.add_paragraph("??")
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
