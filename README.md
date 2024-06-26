# tile_img_on_xl

[> 日本語版](https://github.com/satamame/tile_img_on_xl/blob/master/README_ja.md)

This program slices a picture into small pieces and tile them on an Excel sheet. This is needed because Excel automatically reduces the resolution when you add a big picture.

- For Windows only.
- Generally speaking, adding a large image on an Excel sheet is not a good idea, but it can't be a reason to allow reducing the resolution automatically.

## Environment

This program is confirmed working in below environment.

- Windows 10 (2004)
- Python 3.12
- Microsoft Excel 2016

## Setup

1. In the "tile_img_on_xl" directory, setup virtual environment.
    ```
    > python -m venv .venv
    > .venv\Scripts\activate
    (.venv) > pip install -r requirements.txt
    ```
    **notes**  
    The above procedure installs [pywin32](https://pypi.org/project/pywin32/) via PyPI. However, you might have to download the installer via GitHub. It depends on your environment. See [Release 306 - GitHub](https://github.com/mhammond/pywin32/releases/tag/b306) for details.
1. Save a copy of conf_sample.py as conf.py.

## Configuration

In "conf.py", you can change below variables.

- `max_w`
    - Max width of each piece of sliced picture in px.  
    0 means there's no limitation of width.
- `max_h`
    - Max height of each piece of sliced picture in px.  
    0 means there's no limitation of height.
- `scale`
    - Relative size to the original size that Excel defines.  
    "Original size" depends on Excel's "pixels per point" definition.
- `grouping`
    - Whether to group the pieces after tiling on the sheet.  
    If the image wasn't big enough to slice, it isn't grouped regardless of this setting.

## How to use

1. Open an Excel Workbook, select a Cell which you want to add a big picture.
1. Drag and drop the big picture file to "tile_img_on_xl.bat".
3. That's it. It slices the picture into pieces and tile them on the sheet, then group them if `grouping` is `True`.  
In case the original image wasn't big enough, it isn't sliced or grouped.
