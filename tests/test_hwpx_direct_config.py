from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

import hwpx_direct
import hwpx_hierarchy


def _write_fake_helper(skill_dir: Path) -> None:
    scripts_dir = skill_dir / "scripts"
    scripts_dir.mkdir(parents=True)
    (scripts_dir / "hwpx_helpers.py").write_text(
        "\n".join(
            [
                "NS_DECL = 'xmlns:hp=\"hp\" xmlns:hs=\"hs\"'",
                "_next = 0",
                "def reset_id():",
                "    global _next",
                "    _next = 0",
                "def next_id():",
                "    global _next",
                "    _next += 1",
                "    return str(_next)",
                "def xml_escape(value):",
                "    return str(value).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')",
                "def validate_header_for_government(path):",
                "    return None",
                "def extract_secpr_and_colpr(path):",
                "    return '<secpr/>', '<colpr/>'",
                "def make_first_para(secpr, colpr):",
                "    return '<first>' + secpr + colpr + '</first>'",
                "def make_cover_page(title, subtitle, date):",
                "    return ['<cover>' + xml_escape(title) + '</cover>']",
                "def make_section_bar(number, title):",
                "    return '<section>' + xml_escape(number) + ':' + xml_escape(title) + '</section>'",
                "def make_empty_line():",
                "    return '<empty/>'",
                "def make_text_para(text, charpr='38', parapr='4'):",
                "    return '<text>' + xml_escape(text) + '</text>'",
                "def make_body_para(label, text):",
                "    return '<body>' + xml_escape(label) + ':' + xml_escape(text) + '</body>'",
                "",
            ]
        ),
        encoding="utf-8",
    )


def test_resolve_skill_dir_uses_env_override(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.setenv("HWPX_SKILL_DIR", str(tmp_path))

    assert hwpx_direct.resolve_skill_dir() == tmp_path.resolve()


def test_load_hwpx_helpers_missing_path_is_actionable(tmp_path: Path) -> None:
    error: ImportError | None = None
    try:
        hwpx_direct.load_hwpx_helpers(tmp_path)
    except ImportError as exc:
        error = exc

    assert error is not None
    message = str(error)
    assert str(tmp_path / "scripts" / "hwpx_helpers.py") in message
    assert "--skill-dir" in message
    assert "HWPX_SKILL_DIR" in message


def test_write_header_with_hierarchy_reports_missing_para_properties(tmp_path: Path) -> None:
    # Given
    header_path = tmp_path / "header.xml"
    output_path = tmp_path / "out-header.xml"
    header_path.write_text("<hh:head></hh:head>", encoding="utf-8")

    # When
    error: hwpx_hierarchy.MissingHeaderParaPropertiesError | None = None
    try:
        hwpx_hierarchy.write_header_with_hierarchy(header_path, output_path)
    except hwpx_hierarchy.MissingHeaderParaPropertiesError as exc:
        error = exc

    # Then
    assert error is not None
    assert error.header_path == header_path
    assert "header.xml에서 hh:paraProperties를 찾지 못함" in str(error)
    assert not output_path.exists()


def test_main_forwards_skill_dir_option(tmp_path: Path, capsys) -> None:
    markdown_path = tmp_path / "sample.md"
    output_path = tmp_path / "sample.hwpx"
    skill_dir = tmp_path / "skill"
    markdown_path.write_text("# 제목\n", encoding="utf-8")

    with patch.object(hwpx_direct, "convert_markdown", return_value=output_path) as convert_markdown:
        exit_code = hwpx_direct.main([str(markdown_path), "-o", str(output_path), "--skill-dir", str(skill_dir)])

    captured = capsys.readouterr()
    assert exit_code == 0
    convert_markdown.assert_called_once_with(str(markdown_path), output=str(output_path), skill_dir=str(skill_dir))
    assert str(output_path) in captured.out


def test_build_section_xml_uses_configured_helper_and_renders_table_and_list(tmp_path: Path) -> None:
    skill_dir = tmp_path / "skill"
    _write_fake_helper(skill_dir)
    source = tmp_path / "sample.md"
    section_path = tmp_path / "section0.xml"
    source.write_text("# 제목\n", encoding="utf-8")
    blocks = [
        {"type": "h", "level": 1, "text": "문서 제목"},
        {"type": "h", "level": 2, "text": "1. 개요"},
        {"type": "table", "header": ["구분", "내용"], "rows": [["A", "첫째"]]},
        {"type": "li", "depth": 0, "marker": "○", "content": "목록", "text": "○ 목록"},
    ]

    with patch.object(hwpx_direct, "detect_and_parse", return_value=blocks):
        title = hwpx_direct.build_section_xml(source, section_path, skill_dir=skill_dir)

    xml = section_path.read_text(encoding="utf-8")
    assert title == "문서 제목"
    assert "<hp:tbl" in xml
    assert "구분" in xml
    assert "첫째" in xml
    assert "○ " in xml
    assert "목록" in xml


def test_build_section_xml_renders_merged_table_span_metadata(tmp_path: Path) -> None:
    # Given
    skill_dir = tmp_path / "skill"
    _write_fake_helper(skill_dir)
    source = tmp_path / "sample.md"
    section_path = tmp_path / "section0.xml"
    source.write_text("# 제목\n", encoding="utf-8")
    blocks = [
        {
            "type": "table",
            "header": ["항목", "", "예산액"],
            "rows": [["강사료", "산출내역", "400,000원"]],
            "merged_cells": [[0, 0, 1, 2]],
        }
    ]

    # When
    with patch.object(hwpx_direct, "detect_and_parse", return_value=blocks):
        hwpx_direct.build_section_xml(source, section_path, skill_dir=skill_dir)

    # Then
    xml = section_path.read_text(encoding="utf-8")
    assert '<hp:cellSpan colSpan="2" rowSpan="1"/>' in xml


def test_build_section_xml_uses_section_content_width_for_tables(tmp_path: Path) -> None:
    # Given
    skill_dir = tmp_path / "skill"
    _write_fake_helper(skill_dir)
    helper_path = skill_dir / "scripts" / "hwpx_helpers.py"
    helper_text = helper_path.read_text(encoding="utf-8")
    helper_path.write_text(
        helper_text.replace(
            "    return '<secpr/>', '<colpr/>'",
            (
                "    return '<hp:secPr xmlns:hp=\"hp\"><hp:pagePr width=\"30000\" height=\"40000\">"
                '<hp:margin left="3000" right="2000" top="0" bottom="0" header="0" footer="0" gutter="1000"/>'
                "</hp:pagePr></hp:secPr>', '<colpr/>'"
            ),
        ),
        encoding="utf-8",
    )
    source = tmp_path / "sample.md"
    section_path = tmp_path / "section0.xml"
    source.write_text("# 제목\n", encoding="utf-8")
    blocks = [{"type": "table", "header": ["용어", "정의"], "rows": [["위탁", "외부 기관에 맡김"]]}]

    # When
    with patch.object(hwpx_direct, "detect_and_parse", return_value=blocks):
        hwpx_direct.build_section_xml(source, section_path, skill_dir=skill_dir)

    # Then
    xml = section_path.read_text(encoding="utf-8")
    assert '<hp:sz width="24000"' in xml
    assert '<hp:cellSz width="7200"' in xml
    assert '<hp:cellSz width="16800"' in xml
