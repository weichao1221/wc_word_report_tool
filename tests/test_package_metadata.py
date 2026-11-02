import re
from pathlib import Path

from wc_word_report_tool import __version__

ROOT = Path(__file__).resolve().parents[1]


def test_declared_versions_are_in_sync():
    pyproject = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    word_module = (
        ROOT / "src/wc_word_report_tool/word.py"
    ).read_text(encoding="utf-8")

    project_version = re.search(
        r'^version = "([^"]+)"$', pyproject, re.MULTILINE,
    ).group(1)

    assert project_version == __version__
    assert f"> 版本：v{__version__}" in readme
    assert f"版本：v{__version__}" in word_module


def test_package_no_longer_depends_on_python_docx():
    pyproject = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    source = "\n".join(
        path.read_text(encoding="utf-8")
        for path in (ROOT / "src").rglob("*.py")
    )

    assert "python-docx" not in pyproject
    assert "lxml>=" in pyproject
    assert "from docx" not in source
    assert "import docx" not in source
