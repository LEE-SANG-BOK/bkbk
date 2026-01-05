#!/usr/bin/env python3
from __future__ import annotations

import argparse
import platform
import shutil
import subprocess
from pathlib import Path


def _run(cmd: list[str]) -> None:
    subprocess.run(cmd, check=True)


def _have_cmd(name: str) -> bool:
    return shutil.which(name) is not None


def _find_app_bundle(app_basename: str) -> Path | None:
    """
    Best-effort macOS app bundle lookup.
    We intentionally do not rely on Spotlight/MD queries to keep this script dependency-free.
    """
    candidates = [
        Path("/Applications") / f"{app_basename}.app",
        Path.home() / "Applications" / f"{app_basename}.app",
    ]
    for p in candidates:
        if p.exists():
            return p
    return None


def _soffice_cmd() -> list[str] | None:
    """
    Resolve LibreOffice 'soffice' binary even when not on PATH.
    """
    p = shutil.which("soffice")
    if p:
        return [p]

    candidates = [
        Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        Path.home() / "Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        if c.exists():
            return [str(c)]
    return None


def _convert_with_soffice(docx: Path, out_pdf: Path) -> None:
    soffice = _soffice_cmd()
    if not soffice:
        raise RuntimeError("soffice not found (install LibreOffice or put soffice in PATH)")

    out_dir = out_pdf.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    # LibreOffice naming is based on input basename.
    _run([*soffice, "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), str(docx)])

    expected = out_dir / (docx.stem + ".pdf")
    if expected.resolve() != out_pdf.resolve():
        if expected.exists() and not out_pdf.exists():
            expected.replace(out_pdf)

    if not out_pdf.exists():
        raise RuntimeError(f"soffice did not produce expected PDF: {out_pdf}")


def _convert_with_word_osascript(docx: Path, out_pdf: Path, *, update_fields: bool) -> None:
    if not _have_cmd("osascript"):
        raise RuntimeError("osascript not found")
    if platform.system() == "Darwin" and not _find_app_bundle("Microsoft Word"):
        raise RuntimeError("Microsoft Word.app not found in /Applications")

    # Use VBA via AppleScript for better compatibility.
    # - Fields.Update
    # - TablesOfContents.Update
    # - ExportAsFixedFormat(PDF)
    vba_update = "ActiveDocument.Fields.Update\n" + (
        "On Error Resume Next\n"  # noqa: ISC003
        "Dim toc As TableOfContents\n"
        "For Each toc In ActiveDocument.TablesOfContents\n"
        "  toc.Update\n"
        "Next\n"
        "On Error GoTo 0\n"
    )

    vba_export = (
        f"ActiveDocument.ExportAsFixedFormat OutputFileName:=\"{out_pdf.as_posix()}\", "
        "ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False"
    )

    # Note: Word requires absolute POSIX paths.
    # AppleScript receives argv: docx_path, update_flag, vba_update, vba_export.
    # Passing VBA strings via argv avoids fragile quote-escaping inside the AppleScript source.
    script_lines = [
        "on run argv",
        "  set docxPath to POSIX file (item 1 of argv)",
        "  set doUpdate to (item 2 of argv) as boolean",
        "  set vbUpdate to (item 3 of argv) as text",
        "  set vbExport to (item 4 of argv) as text",
        "  tell application \"Microsoft Word\"",
        "    activate",
        "    set doc to open docxPath",
        "    if doUpdate then",
        "      do visual basic vbUpdate",
        "    end if",
        "    save doc",
        "    do visual basic vbExport",
        "    close doc saving yes",
        "  end tell",
        "end run",
    ]

    out_pdf.parent.mkdir(parents=True, exist_ok=True)

    cmd = [
        "osascript",
        "-e",
        "\n".join(script_lines),
        str(docx.resolve()),
        "true" if update_fields else "false",
        vba_update.replace("\n", "\r"),
        vba_export,
    ]
    _run(cmd)

    if not out_pdf.exists():
        raise RuntimeError(f"Word did not produce expected PDF: {out_pdf}")


def _convert_with_pages_osascript(docx: Path, out_pdf: Path) -> None:
    if not _have_cmd("osascript"):
        raise RuntimeError("osascript not found")
    if platform.system() == "Darwin" and not _find_app_bundle("Pages"):
        raise RuntimeError("Pages.app not found in /Applications")

    out_pdf.parent.mkdir(parents=True, exist_ok=True)

    # Pages can open DOCX and export to PDF on macOS without Microsoft Word.
    script_lines = [
        "on run argv",
        "  set docxPath to POSIX file (item 1 of argv)",
        "  set outPdf to POSIX file (item 2 of argv)",
        "  tell application \"Pages\"",
        "    activate",
        "    set d to open docxPath",
        "    export d to outPdf as PDF",
        "    close d saving no",
        "  end tell",
        "end run",
    ]

    cmd = [
        "osascript",
        "-e",
        "\n".join(script_lines),
        str(docx.resolve()),
        str(out_pdf.resolve()),
    ]
    _run(cmd)

    if not out_pdf.exists():
        raise RuntimeError(f"Pages did not produce expected PDF: {out_pdf}")


def main() -> None:
    ap = argparse.ArgumentParser(
        description=(
            "Postprocess report.docx -> PDF (best-effort). "
            "Auto mode (macOS): Word -> Pages -> LibreOffice(soffice)."
        )
    )
    ap.add_argument("--docx", required=True, type=Path, help="Input .docx path")
    ap.add_argument("--out-pdf", type=Path, default=None, help="Output .pdf path (default: same dir, same stem)")
    ap.add_argument(
        "--mode",
        choices=["auto", "word", "pages", "soffice"],
        default="auto",
        help="Conversion backend",
    )
    ap.add_argument("--no-update-fields", action="store_true", help="Skip Word field/TOC updates")
    ap.add_argument("--open", action="store_true", help="Open the resulting PDF (macOS: open)")

    args = ap.parse_args()

    docx = args.docx.expanduser().resolve()
    if not docx.exists():
        raise SystemExit(f"docx not found: {docx}")

    out_pdf = args.out_pdf.expanduser().resolve() if args.out_pdf else docx.with_suffix(".pdf")

    mode = args.mode
    update_fields = not args.no_update_fields

    tried: list[str] = []
    last_err: Exception | None = None

    def _try(name: str, fn):
        nonlocal last_err
        tried.append(name)
        try:
            fn()
            return True
        except Exception as e:
            last_err = e
            return False

    def _can_word() -> bool:
        return platform.system() == "Darwin" and _have_cmd("osascript") and bool(_find_app_bundle("Microsoft Word"))

    def _can_pages() -> bool:
        return platform.system() == "Darwin" and _have_cmd("osascript") and bool(_find_app_bundle("Pages"))

    def _can_soffice() -> bool:
        return _soffice_cmd() is not None

    if mode in {"word", "auto"}:
        if _can_word():
            ok = _try("word", lambda: _convert_with_word_osascript(docx, out_pdf, update_fields=update_fields))
            if ok:
                print(f"OK (word) wrote: {out_pdf}")
                mode = "done"
            elif mode == "word":
                raise SystemExit(f"Word conversion failed: {last_err}")
        elif mode == "word":
            raise SystemExit("Word mode requires Microsoft Word.app + osascript (macOS)")

    if mode in {"pages", "auto"}:
        if _can_pages():
            ok = _try("pages", lambda: _convert_with_pages_osascript(docx, out_pdf))
            if ok:
                print(f"OK (pages) wrote: {out_pdf}")
                mode = "done"
            elif mode == "pages":
                raise SystemExit(f"Pages conversion failed: {last_err}")
        elif mode == "pages":
            raise SystemExit("Pages mode requires Pages.app + osascript (macOS)")

    if mode in {"soffice", "auto"}:
        if _can_soffice():
            ok = _try("soffice", lambda: _convert_with_soffice(docx, out_pdf))
            if ok:
                print(f"OK (soffice) wrote: {out_pdf}")
                mode = "done"
            elif mode == "soffice":
                raise SystemExit(f"soffice conversion failed: {last_err}")
        elif mode == "soffice":
            raise SystemExit("soffice mode requires LibreOffice (soffice binary)")

    if not out_pdf.exists():
        raise SystemExit(f"All conversion attempts failed (tried={tried}): {last_err}")

    if args.open:
        if platform.system() == "Darwin" and _have_cmd("open"):
            subprocess.run(["open", str(out_pdf)], check=False)


if __name__ == "__main__":
    main()
