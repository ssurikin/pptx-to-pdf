# -*- coding: utf-8 -*-
"""
Конвертер PPTX -> PDF для Windows с сохранением качества изображений.

ПОЧЕМУ РАНЬШЕ PDF ПОЛУЧАЛСЯ «СЖАТЫМ»
Штатный экспорт PowerPoint в PDF (SaveAs / ExportAsFixedFormat) ВСЕГДА
пережимает растровые изображения (JPEG + понижение эффективного
разрешения). Это внутреннее ограничение PowerPoint: оно не отключается
ни настройками программы, ни реестром. Поэтому презентация на 17 МБ
превращалась в PDF на 1 МБ с испорченными картинками.

КАК РАБОТАЕТ ТЕПЕРЬ
Режим raster (по умолчанию): PowerPoint сам отрисовывает каждый слайд
в PNG с заданным DPI (по умолчанию 300), после чего PNG собираются
в PDF без какого-либо перекодирования (img2pdf). Картинки не пережимаются
в принципе, а внешний вид страниц 1:1 совпадает с PowerPoint.

Альтернативные режимы:
  libreoffice — LibreOffice headless с фильтром UseLosslessCompression:
                векторный (выделяемый) текст + несжатые изображения;
  vector      — штатный экспорт PowerPoint (векторный текст, но картинки
                пережимаются самим PowerPoint; оставлен для совместимости).

Готовые PDF сохраняются в папку `output` рядом со скриптом.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import textwrap
import winreg
from pathlib import Path

SUPPORTED_EXTENSIONS = {".pptx", ".ppt", ".pptm", ".pps", ".ppsx", ".ppsm"}

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIR = SCRIPT_DIR / "output"

# Константы PowerPoint COM
PP_FIXED_FORMAT_TYPE_PDF = 2      # ppFixedFormatTypePDF
PP_FIXED_FORMAT_INTENT_PRINT = 2  # ppFixedFormatIntentPrint

PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"

MODE_TITLES = {
    "raster": "PowerPoint: слайды -> PNG -> PDF без пережатия (максимальное качество)",
    "libreoffice": "LibreOffice: PDF с несжатыми изображениями (векторный текст)",
    "vector": "PowerPoint: штатный PDF-экспорт (векторный текст, НО картинки пережимаются)",
}


def safe_call(func, *args):
    """Вызывает func, глотая любые исключения (для COM-очистки)."""
    try:
        return func(*args)
    except Exception:
        return None


# ----------------------------------------------------------------------------
# Настройки «не сжимать картинки» в PowerPoint (реестр HKCU)
# ----------------------------------------------------------------------------

def set_high_quality_mode(office_version: str | None):
    """Лучше-усилия: отключает сжатие изображений в PowerPoint (реестр HKCU).

    Влияет в первую очередь на сохранение PPTX; на штатный PDF-экспорт
    PowerPoint эти ключи, увы, не действуют — но при открытии презентации
    отключённое сжатие исключает дополнительные потери.
    Возвращает состояние для restore_quality_settings() либо None.
    """
    version = (office_version or "16.0").strip()
    key_path = rf"Software\Microsoft\Office\{version}\PowerPoint\Options"
    try:
        key = winreg.CreateKeyEx(
            winreg.HKEY_CURRENT_USER, key_path, 0,
            winreg.KEY_READ | winreg.KEY_WRITE,
        )
    except OSError:
        return None

    saved = {}
    try:
        for name in ("DisableCompression", "DefaultResolution"):
            try:
                saved[name] = winreg.QueryValueEx(key, name)
            except OSError:
                saved[name] = None
        winreg.SetValueEx(key, "DisableCompression", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "DefaultResolution", 0, winreg.REG_DWORD, 0)  # High fidelity
    finally:
        key.Close()
    return (key_path, saved)


def restore_quality_settings(state) -> None:
    """Возвращает исходные значения реестра после экспорта."""
    if not state:
        return
    key_path, saved = state
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE)
    except OSError:
        return
    try:
        for name, value in saved.items():
            if value is None:
                try:
                    winreg.DeleteValue(key, name)
                except OSError:
                    pass
            else:
                winreg.SetValueEx(key, name, 0, value[1], value[0])
    finally:
        key.Close()


# ----------------------------------------------------------------------------
# Общая обвязка PowerPoint COM (pywin32 -> comtypes)
# ----------------------------------------------------------------------------

def powerpoint_com_available() -> bool:
    """Есть ли библиотеки для работы с PowerPoint через COM."""
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
        return True
    except ImportError:
        pass
    try:
        import comtypes.client  # noqa: F401
        return True
    except ImportError:
        return False


def _run_powerpoint_job(app, pptx_path: Path, job):
    """Открывает презентацию в уже запущенном приложении и выполняет job()."""
    version = None
    try:
        version = app.Version
    except Exception:
        pass

    hq_state = set_high_quality_mode(version)  # до открытия презентации
    presentation = None
    try:
        # Open(FileName, ReadOnly=True, Untitled=False, WithWindow=False)
        presentation = app.Presentations.Open(str(pptx_path), True, False, False)
        return job(presentation)
    finally:
        restore_quality_settings(hq_state)
        if presentation is not None:
            safe_call(presentation.Close)


def with_powerpoint(pptx_path: Path, job):
    """Запускает PowerPoint, открывает презентацию, выполняет job(presentation).

    job возвращает свой результат; презентация и приложение закрываются
    в любом случае. Сначала пробуется pywin32, затем comtypes.
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        pythoncom = None
    else:
        pythoncom.CoInitialize()
        try:
            app = None
            try:
                # DispatchEx — отдельный экземпляр PowerPoint, не трогаем
                # уже открытый у пользователя
                try:
                    app = win32com.client.DispatchEx("PowerPoint.Application")
                except Exception:
                    app = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
                return _run_powerpoint_job(app, pptx_path, job)
            finally:
                if app is not None:
                    safe_call(app.Quit)
        finally:
            pythoncom.CoUninitialize()

    # pywin32 не установлен — пробуем comtypes
    import comtypes.client

    app = comtypes.client.CreateObject("PowerPoint.Application")
    try:
        return _run_powerpoint_job(app, pptx_path, job)
    finally:
        safe_call(app.Quit)


# ----------------------------------------------------------------------------
# Режим raster: слайды -> PNG/JPG -> PDF без перекодирования
# ----------------------------------------------------------------------------

def slide_is_hidden(slide) -> bool:
    try:
        return bool(slide.SlideShowTransition.Hidden)
    except Exception:
        return False


def export_slides_to_images(presentation, out_dir: Path, dpi: int,
                            fmt: str, include_hidden: bool):
    """Экспортирует слайды в PNG (или JPG) заданного DPI.

    Возвращает (список файлов, ширина слайда в пунктах).
    """
    setup = presentation.PageSetup
    slide_w_pt = float(setup.SlideWidth)
    slide_h_pt = float(setup.SlideHeight)
    px_w = max(1, round(slide_w_pt / 72.0 * dpi))
    px_h = max(1, round(slide_h_pt / 72.0 * dpi))

    filter_name = "PNG" if fmt == "png" else "JPG"
    ext = ".png" if fmt == "png" else ".jpg"

    slides = presentation.Slides
    total = slides.Count
    images: list[Path] = []

    index = 0
    for i in range(1, total + 1):
        slide = slides.Item(i)
        if not include_hidden and slide_is_hidden(slide):
            continue
        index += 1
        image_path = out_dir / f"slide_{index:03d}{ext}"
        slide.Export(str(image_path), filter_name, px_w, px_h)
        if not image_path.exists():
            raise RuntimeError(f"PowerPoint не смог экспортировать слайд {i}")
        images.append(image_path)
        print(f"\r    слайд {i}/{total}", end="", flush=True)
    print()

    if not images:
        raise RuntimeError("В презентации нет видимых слайдов для экспорта")
    return images, slide_w_pt


def png_pixel_size(path: Path):
    """Читает размеры PNG из заголовка (без Pillow)."""
    try:
        with open(path, "rb") as f:
            header = f.read(24)
    except OSError:
        return None
    if len(header) < 24 or header[:8] != PNG_SIGNATURE or header[12:16] != b"IHDR":
        return None
    return int.from_bytes(header[16:20], "big"), int.from_bytes(header[20:24], "big")


def _strip_alpha(images: list[Path]) -> None:
    """Убирает альфа-канал из PNG (пиксели не меняются, потерь нет)."""
    try:
        from PIL import Image
    except ImportError:
        raise RuntimeError(
            "PowerPoint экспортировал PNG с альфа-каналом. "
            "Установите Pillow: pip install Pillow"
        ) from None

    for path in images:
        with Image.open(path) as im:
            has_alpha = (im.mode in ("RGBA", "LA")
                         or (im.mode == "P" and "transparency" in im.info))
            if not has_alpha:
                continue
            rgba = im.convert("RGBA")
            background = Image.new("RGB", rgba.size, (255, 255, 255))
            background.paste(rgba, mask=rgba.getchannel("A"))
            background.save(path)


def assemble_pdf(images: list[Path], pdf_path: Path,
                 slide_w_pt: float, dpi: int) -> None:
    """Собирает PDF из изображений без перекодирования (img2pdf).

    Размер страницы = размеру слайда в PowerPoint.
    """
    import img2pdf

    layout_dpi = float(dpi)
    actual = png_pixel_size(images[0])
    if actual:
        real_w, _ = actual
        eff_dpi = real_w / (slide_w_pt / 72.0)
        if eff_dpi < dpi * 0.98:
            print(f"[ВНИМАНИЕ] PowerPoint ограничил разрешение до ~{eff_dpi:.0f} DPI")
        layout_dpi = eff_dpi

    layout = img2pdf.get_fixed_dpi_layout_fun((layout_dpi, layout_dpi))
    try:
        data = img2pdf.convert([str(p) for p in images], layout_fun=layout)
    except Exception as exc:
        if "alpha" not in str(exc).lower():
            raise
        print("[ВНИМАНИЕ] PNG с альфа-каналом; убираю прозрачность (без потери качества)...")
        _strip_alpha(images)
        data = img2pdf.convert([str(p) for p in images], layout_fun=layout)

    pdf_path.write_bytes(data)


def convert_raster(pptx_path: Path, output_dir: Path, dpi: int,
                   fmt: str, include_hidden: bool) -> Path:
    """Максимальное качество: PowerPoint рендерит слайды, PDF собирается без сжатия."""
    try:
        import img2pdf  # noqa: F401
    except ImportError:
        raise RuntimeError("режиму raster нужна библиотека img2pdf (pip install img2pdf)")

    pdf_path = output_dir / (pptx_path.stem + ".pdf")

    with tempfile.TemporaryDirectory(prefix="pptx2pdf_") as tmp_name:
        tmp_dir = Path(tmp_name)
        images, slide_w_pt = with_powerpoint(
            pptx_path,
            lambda pres: export_slides_to_images(pres, tmp_dir, dpi, fmt, include_hidden),
        )
        print(f"    собираю PDF из {len(images)} изображений ({fmt.upper()}, без пережатия)...")
        assemble_pdf(images, pdf_path, slide_w_pt, dpi)

    return pdf_path


# ----------------------------------------------------------------------------
# Режим vector: штатный PDF-экспорт PowerPoint (пережимает картинки!)
# ----------------------------------------------------------------------------

def convert_vector(pptx_path: Path, output_dir: Path) -> Path:
    print("[ВНИМАНИЕ] Штатный PDF-экспорт PowerPoint ВСЕГДА пережимает картинки — "
          "это его внутреннее ограничение. Для портфолио используйте режим по умолчанию (raster).")

    pdf_path = output_dir / (pptx_path.stem + ".pdf")

    def job(pres):
        try:
            pres.ExportAsFixedFormat(
                str(pdf_path), PP_FIXED_FORMAT_TYPE_PDF, PP_FIXED_FORMAT_INTENT_PRINT
            )
        except Exception:
            print("[ВНИМАНИЕ] ExportAsFixedFormat недоступен, пробую SaveAs...")
            pres.SaveAs(str(pdf_path), 32)  # ppSaveAsPDF
        return pdf_path

    return with_powerpoint(pptx_path, job)


# ----------------------------------------------------------------------------
# Режим libreoffice: headless + несжатые изображения
# ----------------------------------------------------------------------------

def find_soffice() -> str | None:
    candidates = [
        shutil.which("soffice"),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return candidate
    return None


def convert_libreoffice(pptx_path: Path, output_dir: Path) -> Path:
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice не найден на компьютере")

    pdf_path = output_dir / (pptx_path.stem + ".pdf")
    pdf_path.unlink(missing_ok=True)

    # Фильтр Impress: потерянное сжатие + без понижения разрешения
    lossless_opts = (
        '{"UseLosslessCompression":{"type":"boolean","value":true},'
        '"ReduceImageResolution":{"type":"boolean","value":false},'
        '"Quality":{"type":"long","value":100}}'
    )

    with tempfile.TemporaryDirectory(prefix="lo_profile_") as profile:
        profile_uri = Path(profile).as_uri()
        common = [soffice, "--headless", "--norestore", "--nolockcheck",
                  f"-env:UserInstallation={profile_uri}"]

        result = subprocess.run(
            common + ["--convert-to", f"pdf:impress_pdf_Export:{lossless_opts}",
                      "--outdir", str(output_dir), str(pptx_path)],
            capture_output=True, text=True, errors="replace",
            timeout=1800, check=False,
        )
        if result.returncode != 0 or not pdf_path.exists():
            # Старые LibreOffice (< 7.4) не понимают JSON-опции фильтра
            result = subprocess.run(
                common + ["--convert-to", "pdf",
                          "--outdir", str(output_dir), str(pptx_path)],
                capture_output=True, text=True, errors="replace",
                timeout=1800, check=False,
            )

    if not pdf_path.exists():
        detail = (result.stderr or result.stdout or "").strip()[:400]
        raise RuntimeError(f"LibreOffice не создал PDF. {detail}")
    return pdf_path


# ----------------------------------------------------------------------------
# Основная логика
# ----------------------------------------------------------------------------

def convert_file(pptx_path: Path, output_dir: Path, mode: str,
                 dpi: int, raster_format: str, include_hidden: bool) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)

    if mode == "auto":
        attempts = ["raster", "libreoffice", "vector"]
    else:
        attempts = [mode] + [m for m in ("raster", "libreoffice", "vector") if m != mode]

    problems = []
    total = len(attempts)
    for i, name in enumerate(attempts, 1):
        label = MODE_TITLES[name] if i == 1 else MODE_TITLES[name] + " (резерв)"
        if name in ("raster", "vector") and not powerpoint_com_available():
            problems.append(f"{name}: нет COM-библиотек (pip install pywin32)")
            continue
        try:
            print(f"[{i}/{total}] {label}")
            if name == "raster":
                return convert_raster(pptx_path, output_dir, dpi, raster_format, include_hidden)
            if name == "libreoffice":
                return convert_libreoffice(pptx_path, output_dir)
            return convert_vector(pptx_path, output_dir)
        except Exception as exc:
            problems.append(f"{name}: {exc}")

    raise RuntimeError(
        "Не удалось конвертировать файл.\n  - " + "\n  - ".join(problems)
    )


def normalize_path(raw: str) -> Path:
    """Убирает кавычки и возвращает абсолютный путь."""
    raw = raw.strip().strip('"').strip("'").strip()
    path = Path(raw)
    if not path.is_absolute():
        path = (Path.cwd() / path).resolve()
    return path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptx_to_pdf",
        description="PPTX -> PDF без сжатия изображений.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""\
            Примеры:
              python pptx_to_pdf.py portfolio.pptx
              python pptx_to_pdf.py portfolio.pptx --dpi 450
              python pptx_to_pdf.py portfolio.pptx --raster-format jpg
              python pptx_to_pdf.py portfolio.pptx --mode vector
              python pptx_to_pdf.py portfolio.pptx --mode libreoffice
        """),
    )
    parser.add_argument(
        "files", nargs="*",
        help="пути к презентациям (.pptx, .ppt, ...); без аргументов скрипт спросит",
    )
    parser.add_argument(
        "--mode", choices=("auto", "raster", "libreoffice", "vector"), default="auto",
        help="режим конвертации (по умолчанию auto: raster при наличии PowerPoint, "
             "иначе libreoffice, иначе vector)",
    )
    parser.add_argument(
        "--dpi", type=int, default=300,
        help="DPI для режима raster, 72-600 (по умолчанию 300)",
    )
    parser.add_argument(
        "--raster-format", choices=("png", "jpg"), default="png",
        help="формат слайдов в режиме raster: png — без потерь (по умолчанию), "
             "jpg — заметно компактнее",
    )
    parser.add_argument(
        "--include-hidden", action="store_true",
        help="включить скрытые слайды (режим raster)",
    )
    parser.add_argument(
        "--output", type=Path, default=None,
        help="папка для PDF (по умолчанию: output/ рядом со скриптом)",
    )
    parser.add_argument(
        "--no-open", action="store_true",
        help="не открывать PDF после конвертации",
    )
    return parser


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)

    dpi = args.dpi
    if dpi < 72 or dpi > 600:
        dpi = min(600, max(72, dpi))
        print(f"[ВНИМАНИЕ] DPI вне диапазона 72-600; использую {dpi}")

    output_dir = Path(args.output).resolve() if args.output else DEFAULT_OUTPUT_DIR

    raw_files = args.files or [input("Введите путь к PPTX-файлу: ")]

    failures = 0
    for raw in raw_files:
        pptx_path = normalize_path(raw)

        if not pptx_path.exists():
            print(f"[ОШИБКА] Файл не найден: {pptx_path}")
            failures += 1
            continue

        if pptx_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            print(
                f"[ОШИБКА] Неподдерживаемый формат: '{pptx_path.suffix}'. "
                f"Ожидается один из: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
            )
            failures += 1
            continue

        print(f"\n=== {pptx_path.name} ===")
        try:
            pdf_path = convert_file(
                pptx_path, output_dir, args.mode, dpi,
                args.raster_format, args.include_hidden,
            )
        except Exception as exc:
            print(f"[ОШИБКА] {exc}")
            failures += 1
            continue

        if not pdf_path.exists():
            print(f"[ОШИБКА] PDF не был создан: {pdf_path}")
            failures += 1
            continue

        size_mb = pdf_path.stat().st_size / (1024 * 1024)
        print(f"[ГОТОВО] PDF сохранён: {pdf_path} ({size_mb:.1f} МБ)")
        if not args.no_open:
            os.startfile(str(pdf_path))

    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())