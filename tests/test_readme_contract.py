from __future__ import annotations

from pathlib import Path

from converter_dispatch import SUPPORTED_EXTENSIONS


README_TEXT = Path("README.md").read_text(encoding="utf-8")


def test_readme_mentions_all_supported_extensions() -> None:
    missing = sorted(extension for extension in SUPPORTED_EXTENSIONS if extension not in README_TEXT)

    assert missing == []


def test_readme_mentions_all_public_cli_options() -> None:
    required_options = [
        "--list-formats",
        "--preflight",
        "--startup-timeout",
        "--insert-end-mark",
        "--kordoc-home",
        "-o",
        "--output-dir",
    ]
    missing = [option for option in required_options if option not in README_TEXT]

    assert missing == []


def test_readme_documents_hwp_requirements_and_no_manual_gui_clicking() -> None:
    required_phrases = ["Windows", "한글(HWP)", "HWP COM", "수동 GUI 클릭"]
    missing = [phrase for phrase in required_phrases if phrase not in README_TEXT]

    assert missing == []


def test_readme_includes_bumpis_syntax_mvp_examples() -> None:
    required_terms = ["범피스", "제목:", "소제목:", "네모:", "원:", "바:", "별:", "당구장:", "주석:", "표:", "시간계획표:"]
    missing = [term for term in required_terms if term not in README_TEXT]

    assert missing == []
