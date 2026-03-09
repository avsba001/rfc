import argparse
import shutil
import subprocess
import sys
from pathlib import Path


SUPPORTED_EXTENSIONS = {".docx"}


def bundled_path(relative: str) -> Path:
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1]))
    return base_dir / relative


def resolve_office_binary(user_path: str | None) -> Path:
    if user_path:
        return Path(user_path)

    candidates = [
        bundled_path("third_party/LibreOffice/program/soffice.com"),
        bundled_path("third_party/LibreOffice/program/soffice.exe"),
        Path(r"C:/Program Files/LibreOffice/program/soffice.com"),
        Path(r"C:/Program Files (x86)/LibreOffice/program/soffice.com"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate

    found = shutil.which("soffice") or shutil.which("soffice.com")
    if found:
        return Path(found)

    raise FileNotFoundError("未找到 LibreOffice (soffice)。请通过 --office-path 指定路径。")


def resolve_magick_binary(user_path: str | None) -> Path:
    if user_path:
        return Path(user_path)

    candidates = [
        bundled_path("third_party/ImageMagick/magick.exe"),
        Path(r"C:/Program Files/ImageMagick-7.1.1-Q16-HDRI/magick.exe"),
        Path(r"C:/Program Files (x86)/ImageMagick-7.1.1-Q16-HDRI/magick.exe"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate

    found = shutil.which("magick") or shutil.which("magick.exe")
    if found:
        return Path(found)

    raise FileNotFoundError("未找到 ImageMagick (magick.exe)。请通过 --magick-path 指定路径。")


def collect_docx_files(input_dir: Path, recursive: bool) -> list[Path]:
    iterator = input_dir.rglob("*") if recursive else input_dir.glob("*")
    return [file for file in iterator if file.is_file() and file.suffix.lower() in SUPPORTED_EXTENSIONS]


def run_command(cmd: list[str]) -> None:
    completed = subprocess.run(cmd, capture_output=True, text=True)
    if completed.returncode != 0:
        stderr = completed.stderr.strip()
        stdout = completed.stdout.strip()
        raise RuntimeError(f"命令执行失败: {' '.join(cmd)}\nstdout:\n{stdout}\nstderr:\n{stderr}")


def convert_docx_to_pdf(docx_file: Path, pdf_output_dir: Path, office_bin: Path) -> Path:
    pdf_output_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        str(office_bin),
        "--headless",
        "--convert-to",
        "pdf:writer_pdf_Export",
        "--outdir",
        str(pdf_output_dir),
        str(docx_file),
    ]
    run_command(cmd)

    generated = pdf_output_dir / f"{docx_file.stem}.pdf"
    if not generated.exists():
        raise FileNotFoundError(f"未找到转换后的 PDF: {generated}")
    return generated


def convert_pdf_to_jpg(pdf_file: Path, image_output_dir: Path, magick_bin: Path, dpi: int, quality: int) -> list[Path]:
    image_output_dir.mkdir(parents=True, exist_ok=True)
    output_pattern = image_output_dir / f"{pdf_file.stem}-%03d.jpg"
    cmd = [
        str(magick_bin),
        "-density",
        str(dpi),
        str(pdf_file),
        "-quality",
        str(quality),
        str(output_pattern),
    ]
    run_command(cmd)

    images = sorted(image_output_dir.glob(f"{pdf_file.stem}-*.jpg"))
    if not images:
        raise FileNotFoundError(f"PDF 转 JPG 失败，没有输出图片: {pdf_file}")
    return images


def main() -> None:
    parser = argparse.ArgumentParser(description="批量 DOCX -> PDF -> JPG 转换工具")
    parser.add_argument("input", type=Path, help="输入目录（包含 docx）")
    parser.add_argument("output", type=Path, help="输出目录")
    parser.add_argument("--recursive", action="store_true", help="递归扫描输入目录")
    parser.add_argument("--office-path", type=str, default=None, help="LibreOffice soffice 路径")
    parser.add_argument("--magick-path", type=str, default=None, help="ImageMagick magick.exe 路径")
    parser.add_argument("--dpi", type=int, default=200, help="JPG 输出 DPI")
    parser.add_argument("--quality", type=int, default=92, help="JPG 质量 (1-100)")
    args = parser.parse_args()

    if not args.input.exists() or not args.input.is_dir():
        raise NotADirectoryError(f"输入目录不存在: {args.input}")

    office_bin = resolve_office_binary(args.office_path)
    magick_bin = resolve_magick_binary(args.magick_path)

    docx_files = collect_docx_files(args.input, args.recursive)
    if not docx_files:
        print("没有找到 docx 文件。")
        return

    pdf_dir = args.output / "pdf"
    jpg_root = args.output / "jpg"

    success = 0
    failed = 0
    for docx in docx_files:
        try:
            pdf_path = convert_docx_to_pdf(docx, pdf_dir, office_bin)
            image_dir = jpg_root / docx.stem
            images = convert_pdf_to_jpg(pdf_path, image_dir, magick_bin, args.dpi, args.quality)
            print(f"[OK] {docx.name} -> {pdf_path.name} -> {len(images)} 张 JPG")
            success += 1
        except Exception as exc:
            failed += 1
            print(f"[FAIL] {docx}: {exc}")

    print(f"完成。成功: {success}, 失败: {failed}")


if __name__ == "__main__":
    main()
