"""
制作者 : 棚橋秀斗(p4uw6xas@s.okayama-u.ac.jp)

PDFファイルの制作者を検出する．
できればソースコードと同じフォルダ内にダウンロードしたフォルダ一式を入れること．
"""

import os
import sys
import glob
import pathlib
from typing import List, Tuple, Optional, Any
import re

# 外部モジュール
# conda, pip等でインストールすること
import numpy as np
import pandas as pd
import pypdf
import pdf2image

# pdf2imageの使用にはpoppler（PDF関連のライブラリ）を別途インストールする必要あり
# インストール方法等の詳細は https://github.com/Belval/pdf2image を参照

# 内容は
# (
#     ファイル名,
#     pypdf.metadata.author,
#     pypdf.metadata.creation_date,
#     pypdf.metadata.modification_date
# )
# を想定
ExtractMetaData = Tuple[str, Optional[str], Optional[str], Optional[str]]

# メタデータで確認する要素
metalist = ["Author", "CreationDate", "ModDate"]


def get_metadata(path_pdf: str) -> ExtractMetaData:
    """
    pdfファイルからメタデータを取得

    Parameter
    _________
    path_pdf:
        メタデータを読み取るpdfファイルパス

    Return
    ______
    extract_metadata:
        pdfのメタデータ
    """
    pdf = pypdf.PdfReader(path_pdf)
    author = pdf.metadata.author
    if author == "":
        author = None
    try:
        creation_date = str(pdf.metadata["/CreationDate"])
    except KeyError:
        print(f"Creation date of file {path_pdf} cannot be read.")
        creation_date = None
    try:
        mod_date = str(pdf.metadata["/ModDate"])
    except KeyError:
        print(f"Mod date of file {path_pdf} cannot be read.")
        mod_date = None

    return (pdf, author, creation_date, mod_date)


def get_list_metadata(list_pdf: List[str]) -> pd.DataFrame:
    """
    メタデータを集約した`DataFrame`を作成

    Parameter
    _________
    list_pdf:
        解析結果のpdfリスト

    Return
    ______
    df_meta:
        メタデータのリスト
    """
    df_meta = pd.DataFrame(index=list_pdf, columns=metalist)  # 空のデータフレーム
    for pp in list_pdf:
        try:
            metadata = get_metadata(pp)
            for i in range(len(metalist)):
                df_meta.at[pp, metalist[i]] = metadata[i + 1]
        except ():
            print(f"{pp} cannot open.")

    # inf を消去する
    for c in df_meta.columns:
        df_meta[c] = df_meta[c].replace([np.inf, -np.inf], np.nan)

    return df_meta


def split_noname(df_meta: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    pdfのメタデータをAuthorの記載の有無で分ける

    Parameter
    _________
    df_meta:
        解析対象のメタデータ群

    Return
    ______
    df_author:
        著者名のあるpdfファイル群
    df_noname:
        著者名の無いpdfファイル群
    """
    df_author = df_meta[df_meta["Author"].notna()]
    df_noname = df_meta[df_meta["Author"].isna()].sort_values("ModDate")
    return (df_author, df_noname)


def check_author(df_meta: pd.DataFrame):
    """
    著者名が重複しているファイルを出力

    Parameter
    _________
    df_meta:
        解析対象のメタデータ群

    Return
    ______
    df_duple:
        著者名が重複しているメタデータ
    """
    authors = df_meta["Author"].value_counts()
    duple_auth = authors[authors >= 2].index
    df_duple = df_meta[df_meta["Author"].isin(duple_auth)].sort_values(
        ["Author", "ModDate"]
    )
    return df_duple


def pdf2png(path_pdf: str, path_png: Optional[str] = None) -> Any:
    """
    pdfファイルを画像として取得

    Parameters
    __________
    path_pdf:
        対象とするpdfファイルパス
    path_png: default None
        保存するpngファイルパス
        Noneなら画像保存はなし

    Return
    ______
    images:
        pdf2imageで取得した画像データ
    """
    images = pdf2image.convert_from_path(path_pdf, fmt="png")

    if path_png is not None:
        try:
            images[0].save(path_png, "png")
        except IndexError:
            print(f"Figure cannot save as png: {path_png}")
            # 注意：
            # popplerがインストールされていない場合はpdf2image
            # で画像へ変換できないためこのエラーがでます
    return images


def gen_path_png(path_origin: str, path_dir: str = "images") -> str:
    """
    pngファイルのファイルパスを生成

    Parameters
    __________
    path_origin:
        元ネタにするファイル名
    path_dir: default "images"
        保存先のディレクトリ

    Return
    ______
    path_png:
        pngファイルパス
    """
    name_png = pathlib.Path(path_origin).with_suffix(".png").name
    path_png = pathlib.Path(path_dir).joinpath(name_png)
    return str(path_png)


def check_as_img(
    df_meta: pd.DataFrame, path_img_container: Optional[str] = None
) -> pd.DataFrame:
    """
    pdfを画像データに変換することで一致しているか確認

    Parameters
    __________
    df_meta:
        メタデータ一覧
    path_img_container: default None
        画像の保存先
        Noneならば画像を保存しない

    Return
    ______
    df_dup:
        重複に関する情報を追加した`df_meta`
    """
    df_dup = df_meta.copy()
    pdf_list = df_dup.index
    dup_img = dict()  # 重複している画像一覧
    uniq_img = dict()  # 重複していない画像一覧
    for p in pdf_list:
        if path_img_container is not None:
            path_png = gen_path_png(p, path_dir=path_img_container)
        else:
            path_png = None

        try:
            img = pdf2png(p, path_png)
            if img in dup_img.values():
                # 既に重複したデータが存在する場合
                tag_img = [k for k, v in dup_img.items() if v == img][0]
                df_dup.at[p, "Image(matched)"] = tag_img
            elif img in uniq_img.values():
                # 新規に重複を発見
                tag_img = len(dup_img)
                dup_p = [k for k, v in uniq_img.items() if v == img][0]
                del uniq_img[dup_p]
                for d in (p, dup_p):
                    df_dup.at[d, "Image(matched)"] = tag_img
                dup_img[tag_img] = img
            else:
                uniq_img[p] = img
        except ():
            print(f"{p} cannot to convert png file.")

    if "Image(matched)" in df_dup.columns:
        df_dup = df_dup.sort_values(["Image(matched)", "ModDate"])
    else:
        # 重複画像が全くない場合
        df_dup = df_dup.sort_values("ModDate")

    return df_dup


def write_excel_sheet(
    df_meta: pd.DataFrame, sheet_name: str, writer: pd.ExcelWriter, path_target
):
    """
    Excelシートへの書き込み

    Parameters
    __________
    df_meta:
        Excelに書き込む情報
    sheet_name:
        Excelシート名
    writer:
        open済みのライター
    path_target:
        今回の探索先フォルダ名
        indexの整理に利用
    """
    df_save = df_meta.copy()
    df_save.index = list(
        map(lambda v: str(pathlib.Path(v).relative_to(path_target)), df_save.index)
    )

    # openpyxl.utils.exceptions.IllegalCharacterErrorへの対策
    # openpyxlで扱えない文字列を除去
    ILLEGALCHARACTERS_RE = re.compile(r"[\000-\010]|[\013-\014]|[\016-\037]")

    def illigal_char_re(data):
        s_data = str(data)
        return ILLEGALCHARACTERS_RE.sub("[REMOVED]", s_data)

    df_save["Author"] = df_save["Author"].map(illigal_char_re)

    df_save.to_excel(writer, sheet_name=sheet_name)


def pdf_copypelnor(path_target="./submisstions"):
    """
    指定されたリスト内のPDFについて制作者の重複（と作成者名無し）を抽出.

    Note
    _______
    PDF2.0規格に準拠した，XMP形式のメタデータには対応していません.
    一応，メタデータ無しのPDFファイルとして検知されます.

    Parameters
    _______
    folda_pdf : str
        捜査対象のpdfファイルが格納されたフォルダのパス

    Outputs
    _______
    metadata.xlsx : Excelファイル
        メタデータの一覧表
    """
    list_pdf = glob.glob(path_target + "/**/*.pdf", recursive=True)
    df_meta = get_list_metadata(list_pdf)
    df_author, df_noname = split_noname(df_meta)
    df_dup_auth = check_author(df_author)
    df_dup_img_auth = check_as_img(df_dup_auth)
    dir_noname = "noname"
    os.mkdir(dir_noname)
    df_dup_img_noname = check_as_img(df_noname, path_img_container=dir_noname)

    with pd.ExcelWriter("metadata.xlsx") as writer:
        write_excel_sheet(df_meta, "メタデータ一覧", writer, path_target)
        write_excel_sheet(df_dup_img_auth, "重複している作成者", writer, path_target)
        write_excel_sheet(df_dup_img_noname, "作成者名なし", writer, path_target)

    print("捜査終了")


if __name__ == "__main__":
    path_target = sys.argv[1]
    pdf_copypelnor(path_target)
