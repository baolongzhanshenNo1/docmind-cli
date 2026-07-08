"""End-to-end tests for CLI build command.

Tests the four-step build flow:
  1. Generate (estimated TOC)
  2. Enforce odd pages
  3. Real page numbers (via LibreOffice PDF)
  4. Re-enforce (final pass)

Uses mocked LibreOffice subprocess calls to avoid requiring
LibreOffice or pymupdf to be installed.
"""
import json
import shutil
from pathlib import Path
from unittest import mock

import pytest
from click.testing import CliRunner

from cli import cli, PROJECT_ROOT


class TestCliBuildE2E:
    """End-to-end tests for the `docmind build` command."""

    @pytest.fixture
    def analysis_json(self, tmp_path):
        """Create a minimal analysis JSON for testing."""
        analysis = {
            "title": "测试论文",
            "author": "测试作者",
            "sections": [
                {"type": "cover", "title": "封面", "content": ""},
                {"type": "statement", "title": "郑重声明", "content": ""},
                {"type": "abstract_cn", "title": "摘要", "content": "摘要内容"},
                {"type": "abstract_en", "title": "ABSTRACT", "content": "Abstract content"},
                {"type": "toc", "title": "目录", "content": ""},
                {"type": "body", "title": "1 绪论", "content": "绪论内容"},
                {"type": "conclusion", "title": "结论", "content": "结论内容"},
                {"type": "references", "title": "参考文献", "content": "参考文献内容"},
                {"type": "acknowledgments", "title": "致谢", "content": "致谢内容"},
            ],
        }
        p = tmp_path / "analysis.json"
        p.write_text(json.dumps(analysis, ensure_ascii=False), encoding="utf-8")
        return p

    @pytest.fixture
    def project_root(self, tmp_path):
        """Set up a minimal project root with templates/."""
        # Copy the thesis template
        tmpl_dir = tmp_path / "templates"
        tmpl_dir.mkdir(parents=True, exist_ok=True)
        src_tmpl = PROJECT_ROOT / "templates" / "thesis.yaml"
        if src_tmpl.exists():
            shutil.copy(src_tmpl, tmpl_dir / "thesis.yaml")
        else:
            # Create minimal template inline
            tmpl_dir.joinpath("thesis.yaml").write_text(
                "name: test\n"
                "page:\n  size: A4\n  margins:\n    top: 2.54\n    bottom: 2.54\n"
                "    left: 3.18\n    right: 3.18\n"
                "fonts:\n  body:\n    name: 宋体\n    size: 12\n"
                "  heading:\n    name: 黑体\n    size: 16\n    bold: true\n"
                "sections:\n"
                "  cover:\n    header: ''\n    footer: ''\n    page_number: none\n    break: nextPage\n"
                "  statement:\n    header: ''\n    footer: ''\n    page_number: none\n    break: nextPage\n"
                "  abstract_cn:\n    header: 摘要\n    footer: centered\n    page_number: { format: roman_upper, start: 1 }\n    break: nextPage\n"
                "  abstract_en:\n    header: ABSTRACT\n    footer: centered\n    page_number: { format: roman_upper, continue: true }\n    break: nextPage\n"
                "  toc:\n    header: 目录\n    footer: centered\n    page_number: { format: roman_upper, continue: true }\n    break: nextPage\n"
                "  body:\n    header: '{chapter_title}'\n    footer: centered\n    page_number: { format: arabic, start: 1 }\n    break: nextPage\n    heading_numbering: true\n"
                "  references:\n    header: 参考文献\n    footer: centered\n    page_number: { format: arabic, continue: true }\n    break: nextPage\n"
                "  conclusion:\n    header: 结论\n    footer: centered\n    page_number: { format: arabic, continue: true }\n    break: nextPage\n"
                "  acknowledgments:\n    header: 致谢\n    footer: centered\n    page_number: { format: arabic, continue: true }\n    break: nextPage\n"
                "  appendix:\n    header: 附录\n    footer: centered\n    page_number: { format: arabic, continue: true }\n    break: nextPage\n",
                encoding="utf-8",
            )
        return tmp_path

    @pytest.fixture
    def fake_libreoffice(self, tmp_path):
        """Create a fake LibreOffice executable."""
        p = tmp_path / "fake_soffice.exe"
        p.write_text("echo fake", encoding="utf-8")
        return p

    def test_build_completes_four_step_flow(
        self, tmp_path, analysis_json, project_root, fake_libreoffice
    ):
        """Full build flow: verify all four steps execute and output file exists."""
        output_docx = tmp_path / "final.docx"

        # Mock subprocess.run (used by enforce_odd_pages and toc_real_pages)
        mock_subprocess = mock.MagicMock()
        mock_subprocess.returncode = 0

        # Mock get_section_start_pages to return reasonable page numbers
        # (odd pages for all sections — no blanks needed)
        mock_start_pages = [1, 2, 3, 5, 7, 9, 11, 13, 15]

        # Mock get_real_page_numbers to avoid pymupdf dependency
        mock_page_map = {str(i): i + 1 for i in range(9)}

        with mock.patch("subprocess.run", return_value=mock_subprocess):
            with mock.patch(
                "enforce.odd_pages.get_section_start_pages",
                return_value=mock_start_pages,
            ):
                with mock.patch(
                    "generator.toc_real_pages.get_real_page_numbers",
                    return_value=mock_page_map,
                ):
                    # Redirect PROJECT_ROOT to tmp_path so intermediate files
                    # go to our temp directory
                    with mock.patch("cli.PROJECT_ROOT", project_root):
                        runner = CliRunner()
                        result = runner.invoke(
                            cli,
                            [
                                "build",
                                str(analysis_json),
                                "-t",
                                "thesis",
                                "-o",
                                str(output_docx),
                                "--libreoffice",
                                str(fake_libreoffice),
                            ],
                        )

        # Check exit code and output
        assert result.exit_code == 0, (
            f"CLI build failed with exit code {result.exit_code}:\n{result.output}"
        )

        # Verify all four steps ran
        output = result.output
        assert "Step 1" in output, f"Step 1 not found in output:\n{output}"
        assert "Step 2" in output, f"Step 2 not found in output:\n{output}"
        assert "Step 3" in output, f"Step 3 not found in output:\n{output}"
        assert "Step 4" in output, f"Step 4 not found in output:\n{output}"

        # Verify final output exists
        assert output_docx.exists(), (
            f"Output file not found at {output_docx}\nCLI output:\n{output}"
        )

        # Verify completion message
        assert "构建完毕" in output, f"Completion message not found:\n{output}"

    def test_build_without_enforce(
        self, tmp_path, analysis_json, project_root, fake_libreoffice
    ):
        """Build with --no-enforce should skip Steps 2-4."""
        output_docx = tmp_path / "final_no_enforce.docx"

        with mock.patch("subprocess.run") as mock_run:
            with mock.patch("cli.PROJECT_ROOT", project_root):
                runner = CliRunner()
                result = runner.invoke(
                    cli,
                    [
                        "build",
                        str(analysis_json),
                        "-t",
                        "thesis",
                        "-o",
                        str(output_docx),
                        "--no-enforce",
                        "--no-real-pages",
                    ],
                )

        assert result.exit_code == 0, (
            f"CLI build (no-enforce) failed with exit code {result.exit_code}:\n{result.output}"
        )

        # Verify only Step 1 ran
        output = result.output
        assert "Step 1" in output
        assert "Step 2" not in output
        assert "Step 3" not in output
        assert "Step 4" not in output

        # Subprocess should not have been called
        mock_run.assert_not_called()

        # Verify output exists
        assert output_docx.exists(), f"Output file not found at {output_docx}"

    def test_build_missing_analysis_file(self):
        """Build with nonexistent analysis JSON should fail."""
        runner = CliRunner()
        result = runner.invoke(
            cli,
            [
                "build",
                "nonexistent.json",
            ],
        )
        assert result.exit_code != 0, "Expected non-zero exit for missing file"

    def test_build_bad_template(self, analysis_json, fake_libreoffice):
        """Build with nonexistent template should fail."""
        runner = CliRunner()
        result = runner.invoke(
            cli,
            [
                "build",
                str(analysis_json),
                "-t",
                "nonexistent_template",
                "-o",
                "out.docx",
                "--libreoffice",
                str(fake_libreoffice),
            ],
        )
        assert result.exit_code != 0, "Expected non-zero exit for bad template"
