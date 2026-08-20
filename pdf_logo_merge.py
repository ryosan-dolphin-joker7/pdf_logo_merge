#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDFの指定領域を、別のPDFの所定位置へベクターのまま合成するスクリプト。

前提:
    - Python 3
    - PyMuPDF（import名は fitz）
    - PyInstaller --onefile でWindows用EXEに変換可能

座標系:
    PyMuPDFの座標（単位: point、原点: ページ左上）を使用する。
    1 point = 1/72 inch である。

重要:
    show_pdf_page() を使用するため、抽出元PDFがベクターデータであれば、
    ラスタライズせずベクターデータのまま出力PDFへ配置される。
"""

from __future__ import annotations

import os
import sys
import uuid
from pathlib import Path
from typing import Final


# PyMuPDFが未導入でも、日本語のエラー表示と終了待機を行えるようにする。
try:
    import fitz  # PyMuPDF
except ImportError as exc:  # pragma: no cover - PyMuPDF未導入時のみ通る
    fitz = None  # type: ignore[assignment]
    FITZ_IMPORT_ERROR: ImportError | None = exc
else:
    FITZ_IMPORT_ERROR = None


# ---------------------------------------------------------------------------
# 設定値
# ---------------------------------------------------------------------------

# EXEまたは本Pythonファイルと同じフォルダに置くファイル名
LOGO_SOURCE_FILE: Final[str] = "logo_source.pdf"
TARGET_FORM_FILE: Final[str] = "target_form.pdf"
OUTPUT_FILE: Final[str] = "output_merged.pdf"

# 抽出元PDFのページ番号（PyMuPDF内部の0始まり。0は1ページ目）
SOURCE_PAGE_INDEX: Final[int] = 0

# 抽出範囲: (左上X, 左上Y, 右下X, 右下Y)
SOURCE_CLIP_COORDS: Final[tuple[float, float, float, float]] = (
    10.0,
    10.0,
    150.0,
    50.0,
)

# 配置範囲: (左上X, 左上Y, 右下X, 右下Y)
DESTINATION_COORDS: Final[tuple[float, float, float, float]] = (
    400.0,
    20.0,
    540.0,
    60.0,
)

# Noneなら全ページへ合成する。
# 指定ページだけに合成したい場合は、1始まりで例のように指定する。
# 例: (1, 3, 5) なら1・3・5ページだけを処理する。
TARGET_PAGES: Final[tuple[int, ...] | None] = None


class PdfMergeError(Exception):
    """利用者に分かりやすく通知するためのアプリケーション例外。"""


def get_application_directory() -> Path:
    """
    実行ファイルが存在するフォルダを返す。

    PyInstallerでEXE化された場合:
        sys.executable は一時展開先ではなく、起動したEXE自身を指す。

    Pythonスクリプトとして実行した場合:
        sys.argv[0]（このスクリプトのパス）の親フォルダを返す。

    これにより、ショートカットや別フォルダから起動しても、カレント
    ディレクトリに依存せず、EXEと同じ場所のPDFを確実に参照できる。
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(sys.argv[0]).resolve().parent


def ensure_distinct_paths(*paths: Path) -> None:
    """入力・出力が同一ファイルを指していないことを確認する。"""
    normalized = [os.path.normcase(str(path.resolve(strict=False))) for path in paths]
    if len(normalized) != len(set(normalized)):
        raise PdfMergeError(
            "入力ファイルと出力ファイルには、それぞれ異なる名前を指定してください。"
        )


def validate_input_file(path: Path, label: str) -> None:
    """入力PDFの存在とファイル種別を確認する。"""
    if not path.exists():
        raise PdfMergeError(f"{label}が見つかりません: {path}")
    if not path.is_file():
        raise PdfMergeError(f"{label}はファイルではありません: {path}")
    if path.suffix.lower() != ".pdf":
        raise PdfMergeError(f"{label}の拡張子がPDFではありません: {path.name}")


def validate_pdf_document(document: "fitz.Document", label: str) -> None:
    """PDFが開ける状態か、ページを持つか、パスワードが必要でないかを確認する。"""
    if document.needs_pass:
        raise PdfMergeError(f"{label}はパスワードで保護されているため処理できません。")
    if document.page_count < 1:
        raise PdfMergeError(f"{label}にページがありません。")


def validate_rectangle(
    rect: "fitz.Rect",
    page_rect: "fitz.Rect",
    label: str,
    page_number: int,
) -> None:
    """矩形が有効であり、対象ページ内に完全に収まることを確認する。"""
    if rect.is_empty or rect.is_infinite or rect.x0 >= rect.x1 or rect.y0 >= rect.y1:
        raise PdfMergeError(f"{label}の座標が不正です: {tuple(rect)}")

    tolerance = 0.001
    is_inside = (
        rect.x0 >= page_rect.x0 - tolerance
        and rect.y0 >= page_rect.y0 - tolerance
        and rect.x1 <= page_rect.x1 + tolerance
        and rect.y1 <= page_rect.y1 + tolerance
    )
    if not is_inside:
        raise PdfMergeError(
            f"{label}が{page_number}ページ目の範囲外です。"
            f" 指定={tuple(rect)}, ページ={tuple(page_rect)}"
        )


def resolve_target_page_indices(page_count: int) -> list[int]:
    """設定値TARGET_PAGESを、PyMuPDF用の0始まりページ番号へ変換する。"""
    if TARGET_PAGES is None:
        return list(range(page_count))

    if not TARGET_PAGES:
        raise PdfMergeError("TARGET_PAGESが空です。Noneまたはページ番号を指定してください。")

    if any(isinstance(number, bool) or not isinstance(number, int) for number in TARGET_PAGES):
        raise PdfMergeError("TARGET_PAGESには1始まりの整数だけを指定してください。")

    if len(TARGET_PAGES) != len(set(TARGET_PAGES)):
        raise PdfMergeError("TARGET_PAGESに同じページ番号が重複しています。")

    invalid_pages = [number for number in TARGET_PAGES if not 1 <= number <= page_count]
    if invalid_pages:
        raise PdfMergeError(
            f"対象PDFのページ数は{page_count}です。範囲外の指定: {invalid_pages}"
        )

    return [number - 1 for number in TARGET_PAGES]


def merge_pdf_area(
    logo_source_path: Path,
    target_form_path: Path,
    output_path: Path,
) -> int:
    """
    抽出元PDFの指定領域を、対象PDFの各ページへ合成する。

    保存途中の失敗で既存出力を壊さないよう、同じフォルダの一時PDFへ
    保存してからos.replace()で出力ファイルへ置き換える。

    Returns:
        合成したページ数。
    """
    if fitz is None:
        raise PdfMergeError(
            "PyMuPDFがインストールされていません。"
            " コマンド『python -m pip install PyMuPDF』を実行してください。"
        ) from FITZ_IMPORT_ERROR

    ensure_distinct_paths(logo_source_path, target_form_path, output_path)
    validate_input_file(logo_source_path, "抽出元PDF")
    validate_input_file(target_form_path, "合成先PDF")

    source_clip = fitz.Rect(SOURCE_CLIP_COORDS)
    destination_rect = fitz.Rect(DESTINATION_COORDS)

    # 同一フォルダに一時ファイルを作ることで、最終置換を安全に行う。
    temporary_output = output_path.with_name(
        f".{output_path.stem}_{uuid.uuid4().hex}.tmp.pdf"
    )

    try:
        # withを使い、成功・失敗にかかわらずPDFを確実に閉じる。
        with fitz.open(logo_source_path) as logo_document, fitz.open(
            target_form_path
        ) as target_document:
            validate_pdf_document(logo_document, "抽出元PDF")
            validate_pdf_document(target_document, "合成先PDF")

            if SOURCE_PAGE_INDEX >= logo_document.page_count:
                raise PdfMergeError(
                    f"抽出元PDFには{logo_document.page_count}ページしかありません。"
                )

            source_page = logo_document.load_page(SOURCE_PAGE_INDEX)
            validate_rectangle(
                source_clip,
                source_page.rect,
                "切り取り範囲",
                SOURCE_PAGE_INDEX + 1,
            )

            page_indices = resolve_target_page_indices(target_document.page_count)

            # 途中まで合成してから座標エラーになることを避けるため、
            # 先に全対象ページの配置範囲を検証する。
            for page_index in page_indices:
                target_page = target_document.load_page(page_index)
                validate_rectangle(
                    destination_rect,
                    target_page.rect,
                    "貼り付け範囲",
                    page_index + 1,
                )

            if source_page.rotation != 0:
                print(
                    f"注意: 抽出元PDFの1ページ目には"
                    f"{source_page.rotation}度の回転情報があります。"
                )

            # show_pdf_page()はページ内容をForm XObjectとして配置するため、
            # 元がベクターならベクターのまま保持される。
            for page_index in page_indices:
                target_page = target_document.load_page(page_index)
                target_page.show_pdf_page(
                    destination_rect,
                    logo_document,
                    pno=SOURCE_PAGE_INDEX,
                    clip=source_clip,
                    keep_proportion=False,  # 指定矩形へ正確にフィットさせる
                    overlay=True,  # 帳票の既存内容より前面に配置する
                    rotate=0,
                )
                print(
                    f"処理中: {page_index + 1}/{target_document.page_count}ページ目へ"
                    "合成しました。"
                )

            # garbage=4: 未使用オブジェクト等を整理
            # deflate=True: 圧縮可能なストリームを圧縮
            # clean=True: ページ内容ストリームを整理
            target_document.save(
                temporary_output,
                garbage=4,
                deflate=True,
                clean=True,
            )

        if not temporary_output.exists() or temporary_output.stat().st_size == 0:
            raise PdfMergeError("出力PDFの一時保存に失敗しました。")

        # 出力が既に存在する場合も、完成した一時PDFで安全に置き換える。
        os.replace(temporary_output, output_path)
        return len(page_indices)

    finally:
        # エラー時に一時ファイルだけが残らないよう後片付けする。
        try:
            if temporary_output.exists():
                temporary_output.unlink()
        except OSError:
            # 本来の処理エラーを、一時ファイル削除エラーで上書きしない。
            pass


def run() -> int:
    """パスを組み立てて処理を実行し、プロセスの終了コードを返す。"""
    print("PDFエリア合成処理を開始します。")

    application_directory = get_application_directory()
    logo_source_path = application_directory / LOGO_SOURCE_FILE
    target_form_path = application_directory / TARGET_FORM_FILE
    output_path = application_directory / OUTPUT_FILE

    print(f"実行フォルダ: {application_directory}")
    print(f"抽出元PDF  : {logo_source_path.name}")
    print(f"合成先PDF  : {target_form_path.name}")
    print(f"出力PDF    : {output_path.name}")

    try:
        merged_page_count = merge_pdf_area(
            logo_source_path=logo_source_path,
            target_form_path=target_form_path,
            output_path=output_path,
        )
    except PdfMergeError as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 1
    except PermissionError as exc:
        print(
            "エラー: ファイルへアクセスできません。PDFが別のアプリで開かれて"
            "いないか、フォルダに書き込み権限があるか確認してください。",
            file=sys.stderr,
        )
        print(f"詳細: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        # 破損PDF、PyMuPDF内部エラー、ディスク容量不足などの想定外エラーも
        # コンソールに残し、無言で終了しないようにする。
        print(
            f"予期しないエラーが発生しました: {type(exc).__name__}: {exc}",
            file=sys.stderr,
        )
        return 1

    print(f"完了: {merged_page_count}ページへ合成しました。")
    print(f"出力ファイル: {output_path}")
    return 0


def wait_for_exit() -> None:
    """EXEのコンソールがすぐ閉じないよう、最後にEnter入力を待つ。"""
    try:
        input("Enterキーを押して終了してください...")
    except (EOFError, KeyboardInterrupt):
        # 入力を受け取れない環境やCtrl+Cでも、異常なスタック表示を出さない。
        pass


if __name__ == "__main__":
    exit_code = 1
    try:
        exit_code = run()
    finally:
        # 成功時・エラー時のどちらでも必ず待機する。
        wait_for_exit()

    raise SystemExit(exit_code)
