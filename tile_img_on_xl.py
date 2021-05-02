import math
import sys
from collections import namedtuple
from pathlib import Path

import win32com.client
from PIL import Image

import conf

Size = namedtuple('Size', 'w h')
Rect = namedtuple('Rect', 'l t r b')


def get_rects(src_size):
    """切り出す矩形のリストを作る

    Parameters
    ----------
    src_size : Size
        切り出す前のサイズ

    Returns
    -------
    rects : list
        切り出す矩形のリスト
    Size
        横方向と縦方向の分割数
    """
    # 横方向の分割数
    if conf.max_w:
        col_cnt = math.ceil(src_size.w / conf.max_w)
    else:
        col_cnt = 1

    # 縦方向の分割数
    if conf.max_h:
        row_cnt = math.ceil(src_size.h / conf.max_h)
    else:
        row_cnt = 1

    # 各ピースの幅と高さ
    unit_w = math.ceil(src_size.w / col_cnt)
    unit_h = math.ceil(src_size.h / row_cnt)

    # 切り出しのための矩形を得る
    rects = []
    for r in range(row_cnt):
        for c in range(col_cnt):
            top = unit_h * r
            left = unit_w * c
            bottom = min(unit_h * (r + 1), src_size.h)
            right = min(unit_w * (c + 1), src_size.w)
            rects.append(Rect(left, top, right, bottom))

    return rects, Size(col_cnt, row_cnt)


def save_sliced_imgs(src_img, rects):
    """画像をスライスして各ピースを保存する

    Parameters
    ----------
    src_img : Image
        元画像
    rects : list
        切り出しに使う矩形のリスト

    Returns
    -------
    img_paths : list
        保存した各ピースのパスのリスト
    """
    temp_dir = Path(__file__).parent.resolve() / 'temp'

    # temp_dir がディレクトリなら空にする
    if temp_dir.is_dir():
        for f in temp_dir.iterdir():
            f.unlink()
    # temp_dir がディレクトリでなければ作る
    else:
        if temp_dir.exists():
            temp_dir.unlink()
        temp_dir.mkdir()

    # 切り出して保存する
    img_paths = []
    for i, rect in enumerate(rects):
        img_paths.append(file_path := temp_dir / f'{i:04}.png')
        src_img.crop(rect).save(file_path)

    return img_paths


def tile_imgs_on_xl(img_paths, counts):
    """画像のピースを Excel の ActiveWorkbook に敷き詰める

    Parameters
    ----------
    img_paths : list
        画像のピースのパスのリスト
    counts : Size
        横方向と縦方向のピース数

    Returns
    -------
    workbook : Workbook
        画像を敷き詰めた Workbook
    """
    # Excel アプリケーションのオブジェクトを取得
    try:
        xl_app = win32com.client.GetObject(Class='Excel.Application')
    except Exception:
        print('Terminated. No Excel application.')
        raise

    # アクティブな Workbook, Sheet, Cell を取得
    workbook = xl_app.ActiveWorkbook
    sheet = xl_app.ActiveSheet
    cell = xl_app.ActiveCell
    if not (workbook and sheet and cell):
        print('Terminated.')
        raise Exception('No active Workbook, Sheet, or Cell.')

    # 画像をファイルから読み込み敷き詰める
    path_iter = iter(img_paths)
    y = cell.Top
    for row in range(counts.h):
        x = cell.Left
        for col in range(counts.w):
            # 画像のピースをシートに追加する
            img_path = next(path_iter)
            sheet.Shapes.AddPicture(
                Filename=img_path,
                LinkToFile=False,
                SaveWithDocument=True,
                Left=x,
                Top=y,
                Width=-1,
                Height=-1
            )

            # 追加した画像オブジェクトを取得する
            shape_count = len(sheet.Shapes)
            last_shape = sheet.Shapes[shape_count - 1]

            # 倍率を適用する
            last_shape.ScaleWidth(conf.scale, True)
            last_shape.ScaleHeight(conf.scale, True)

            x = last_shape.Left + last_shape.Width
        y = last_shape.Top + last_shape.Height

    return workbook


def main(img_path):
    # コマンド引数から画像ファイルと画素数を取得する
    img = Image.open(img_path)
    img_size = Size(*img.size)

    # 切り出す矩形のリストを作る
    rects, counts = get_rects(img_size)

    # 画像をスライスして各ピースを保存する
    img_paths = save_sliced_imgs(img, rects)

    # 保存したピースを Excel に貼り付ける
    workbook = tile_imgs_on_xl(img_paths, counts)

    # 設定によりグループ化する
    if conf.grouping:
        sheet = workbook.ActiveSheet
        shape_count = len(sheet.Shapes)
        shp_names = []
        for i in range(len(rects)):
            shp_names.append(sheet.Shapes[shape_count - i - 1].Name)
        sheet.Shapes.Range(shp_names).Group()

    # 貼り付けた Workbook を前面に表示する
    shell = win32com.client.Dispatch('WScript.Shell')
    shell.AppActivate(workbook.Name)


if __name__ == '__main__':
    if len(sys.argv) > 1:
        main(sys.argv[1])
    else:
        raise Exception('Terminated. No argument.')
