# tile_img_on_xl

[> English version](https://github.com/satamame/tile_img_on_xl/blob/master/README.md)

画像を分割して Excel のシートに貼り付けるプログラムです。大きな画像を Excel のシートに貼る時に、解像度が勝手に低くなる問題を回避します。

- Windows のみに対応します。
- 大きな画像を Excel のシートに貼るという使い方もどうかと思いますが、だからと言って勝手に解像度が下がるというのも変な話です。

## 動作環境

以下の環境で動作確認済みです。

- Windows 10 (2004)
- Python 3.8
- Microsoft Excel 2016

## 準備

1. tile_img_on_xl フォルダで仮想環境を構築します。
    ```
    > python -m venv .venv
    > .venv\Scripts\activate
    (.venv) > pip install -r requirements.txt
    ```
    **注意**  
    上記の手順では、PyPI から [pywin32](https://pypi.org/project/pywin32/) をインストールしますが、環境によっては GitHub からインストーラをダウンロードする必要があるかも知れません。詳しくは [Release 300 - GitHub](https://github.com/mhammond/pywin32/releases/tag/b300) を御覧ください。
1. conf_sample.py をコピーして conf.py にリネームします。

## 設定

conf.py を編集して以下の設定ができます。

- `max_w`
    - 分割時の各ピースの幅の最大値を px で指定します。  
    0 を指定した場合は横方向に分割しません。
- `max_h`
    - 分割時の各ピースの高さの最大値を px で指定します。  
    0 を指定した場合は縦方向に分割しません。
- `scale`
    - Excel により定義される「元のサイズ」に対する縮尺。  
    「元のサイズ」は、Excel が「1ポイントあたりの px 数」をどう定義しているかによります。
- `grouping`
    - 分割したピースを貼り付けた後にグループ化するかどうか。  
    元の画像が小さくて分割されなかった場合は、設定に関係なくグループ化しません。

## 使い方

1. Excel の Workbook を開いて、大きな画像を貼り付けたいセルを選択します。
1. 大きな画像のファイルを "tile_img_on_xl.bat" にドラッグアンドドロップします。
1. これで、画像が分割されシートに並べられます。`grouping` が `True` ならグループ化されます。  
※元の画像が十分な大きさでない場合は、分割やグループ化はされません。
