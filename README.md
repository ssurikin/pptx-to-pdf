# PPTX → PDF конвертер (без сжатия картинок)

Конвертер презентаций PowerPoint в PDF для Windows с **сохранением оригинального качества изображений**.

## Почему обычные конвертеры портят картинки

Штатный экспорт PowerPoint в PDF (`SaveAs` / `ExportAsFixedFormat`) **всегда** пережимает растровые изображения: JPEG-компрессия + понижение разрешения. Это внутреннее ограничение PowerPoint, которое не отключается ни настройками программы, ни реестром. Поэтому презентация на 17 МБ превращается в PDF на 1 МБ с мыльными картинками.

## Как этот конвертер решает проблему

По умолчанию (режим `raster`): PowerPoint отрисовывает каждый слайд в PNG с высоким разрешением (по умолчанию 300 DPI, можно поднять до 600), после чего PNG собираются в PDF **без какого-либо перекодирования** (библиотека `img2pdf`). Пиксели изображений не пережимаются ни разу, а внешний вид страниц 1:1 совпадает с тем, что показывает PowerPoint.

## Использование

### Простой способ

1. Запустите `convert.bat` двойным кликом.
2. Введите путь к PPTX-файлу (или просто перетащите файл в окно консоли).
   - Также можно перетащить PPTX-файл прямо **на значок** `convert.bat` — путь подхватится автоматически.
3. PDF появится в папке `output` рядом со скриптом и сразу откроется.

### Через командную строку

```bash
python pptx_to_pdf.py portfolio.pptx              # raster 300 DPI (по умолчанию)
python pptx_to_pdf.py portfolio.pptx --dpi 450    # выше разрешение
python pptx_to_pdf.py portfolio.pptx --raster-format jpg   # компактнее, с потерями
python pptx_to_pdf.py portfolio.pptx --mode libreoffice   # через LibreOffice
python pptx_to_pdf.py portfolio.pptx --mode vector        # штатный экспорт PP (сжимает!)
python pptx_to_pdf.py a.pptx b.pptx c.pptx        # пакетная конвертация
python pptx_to_pdf.py --help                      # все опции
```

## Режимы конвертации

| Режим | Как работает | Качество картинок | Текст |
|---|---|---|---|
| `raster` (по умолчанию) | PowerPoint рендерит слайды в PNG → PDF без перекодирования | **Без потерь** | Растровый (не выделяется) |
| `libreoffice` | LibreOffice headless с фильтром `UseLosslessCompression` | Без потерь (lossless) | Векторный (выделяется) |
| `vector` | Штатный PDF-экспорт PowerPoint | **Сжимается PowerPoint'ом!** | Векторный |

Режим `auto` (по умолчанию) пробует `raster` → `libreoffice` → `vector`, пока что-то не сработает.

## Установка

Требуется Python 3.10+.

```bash
pip install -r requirements.txt
```

- `pywin32` — работа с PowerPoint через COM
- `img2pdf` — сборка PDF из изображений без перекодирования
- `Pillow` — на случай PNG с альфа-каналом

Нужен установленный PowerPoint (для режимов `raster`/`vector`) или LibreOffice (для `libreoffice`).

## Файлы проекта

- `convert.bat` — запуск двойным кликом
- `pptx_to_pdf.py` — основная логика конвертации
- `requirements.txt` — зависимости
- `output/` — сюда сохраняются готовые PDF

## Лицензия

[MIT](LICENSE)