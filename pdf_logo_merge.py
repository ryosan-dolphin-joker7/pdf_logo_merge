#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
請求書・封筒宛名PDF作成ツール（GUI版）

複数ページの請求書PDFから各ページの宛名領域を抽出し、
「封筒宛名ページ → 元の請求書ページ」の順番で1つのPDFへまとめる。
Windows用EXE（PyInstaller --onefile --windowed）での利用を想定している。

主な機能:
    - 請求書PDF、封筒テンプレートPDF、出力先を選択
    - 見本の請求書ページを表示し、宛名範囲をマウスドラッグで指定
    - 封筒テンプレートを表示し、宛名の印刷位置をマウスドラッグで指定
    - 全ページまたは「1,3-5」のように処理する請求書ページを指定
    - 請求書の各ページから同じ座標の宛名を抽出
    - ユーザーごとに封筒宛名ページと請求書ページを交互に出力
    - 宛名範囲のプレビュー
    - 入力・座標・ページ範囲・暗号化PDF・保存先を事前検証
    - 一時保存後に出力を置き換え、保存失敗による既存PDF破損を防止

セキュリティ方針:
    - PDF拡張子だけでなくファイルヘッダーと実際の文書形式も確認する
    - 過大なファイル、ページ数、プレビュー画像による資源枯渇を防止する
    - 入力PDFの上書き、ハードリンク経由の同一ファイル指定を防止する
    - shell / subprocess / 外部通信は使用しない
    - 出力はランダム名の一時PDFへ保存し、成功時だけアトミックに置換する

座標系:
    PyMuPDFの座標（単位: point、原点: ページ左上）を使用する。
    1 point = 1/72 inch、1 mm ≒ 2.83465 point。

重要:
    show_pdf_page() と insert_pdf() を使用するため、元データがベクターなら、
    ラスタライズせずベクターのまま出力PDFへ配置される。
    プレビュー表示だけは画面確認用に一時的にラスタライズする。
"""

from __future__ import annotations

import base64
import math
import os
import sys
import threading
import uuid
from pathlib import Path
from typing import Callable, Final

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


# 公式パッケージ名で読み込み、同名に近い別パッケージとの取り違えを防ぐ。
# PyMuPDFが未導入でも、GUIで分かりやすいエラーを表示できるようにする。
try:
    import pymupdf as fitz
except ImportError as exc:  # pragma: no cover - PyMuPDF未導入時だけ通る
    fitz = None  # type: ignore[assignment]
    FITZ_IMPORT_ERROR: ImportError | None = exc
else:
    FITZ_IMPORT_ERROR = None


APP_TITLE: Final[str] = "請求書・封筒宛名PDF作成ツール"
PDF_FILE_TYPES: Final[tuple[tuple[str, str], ...]] = (
    ("PDFファイル", "*.pdf"),
    ("すべてのファイル", "*.*"),
)

# 前回の固定設定をGUIの初期値として維持する。
DEFAULT_ADDRESS_COORDS: Final[tuple[float, float, float, float]] = (
    10.0,
    10.0,
    150.0,
    50.0,
)
DEFAULT_ENVELOPE_DESTINATION_COORDS: Final[
    tuple[float, float, float, float]
] = (
    400.0,
    20.0,
    540.0,
    60.0,
)

# 不正・破損PDFや誤操作によるメモリ／ディスクの過剰消費を抑える上限。
# 通常の帳票用途には十分余裕を持たせている。必要なら運用に合わせて変更できる。
MAX_INPUT_FILE_BYTES: Final[int] = 1 * 1024 * 1024 * 1024  # 1 GiB
MAX_PDF_PAGES: Final[int] = 20_000
MAX_PAGE_SPECIFICATION_LENGTH: Final[int] = 10_000
MAX_PAGE_DIMENSION_POINTS: Final[float] = 200_000.0
MAX_PREVIEW_PIXELS: Final[int] = 8_000_000
PREVIEW_MAX_WIDTH: Final[int] = 1_000
PREVIEW_MAX_HEIGHT: Final[int] = 650

ProgressCallback = Callable[[str], None]


class PdfMergeError(Exception):
    """画面へそのまま表示できる、利用者向けの処理エラー。"""


def get_application_directory() -> Path:
    """Python実行時はスクリプト、EXE実行時はEXEがあるフォルダーを返す。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(sys.argv[0]).resolve().parent


def ensure_distinct_paths(*paths: Path) -> None:
    """請求書・封筒テンプレート・出力先が同じでないことを確認する。"""
    normalized = [os.path.normcase(str(path.resolve(strict=False))) for path in paths]
    if len(normalized) != len(set(normalized)):
        raise PdfMergeError(
            "請求書PDF、封筒テンプレートPDF、出力PDFには"
            "異なるファイルを指定してください。"
        )

    # 表記上のパスが違っても、ハードリンク等で実体が同じ場合を検出する。
    existing_paths = [path for path in paths if path.exists()]
    for index, left_path in enumerate(existing_paths):
        for right_path in existing_paths[index + 1 :]:
            try:
                if os.path.samefile(left_path, right_path):
                    raise PdfMergeError(
                        "請求書PDF、封筒テンプレートPDF、出力PDFには"
                        "実体の異なるファイルを指定してください。"
                    )
            except OSError:
                # 権限や一時的なファイル状態は、後続のopen/saveで詳細に扱う。
                continue


def validate_input_file(path: Path, label: str) -> None:
    """入力ファイルの存在・サイズ・拡張子・PDFヘッダーを確認する。"""
    if not path.exists():
        raise PdfMergeError(f"{label}が見つかりません。\n{path}")
    if not path.is_file():
        raise PdfMergeError(f"{label}はファイルではありません。\n{path}")
    if path.suffix.lower() != ".pdf":
        raise PdfMergeError(f"{label}にはPDFファイルを指定してください。\n{path}")
    try:
        file_size = path.stat().st_size
    except OSError as exc:
        raise PdfMergeError(f"{label}の情報を取得できません。\n{path}") from exc
    if file_size <= 0:
        raise PdfMergeError(f"{label}は空のファイルです。\n{path}")
    if file_size > MAX_INPUT_FILE_BYTES:
        raise PdfMergeError(
            f"{label}が安全上の上限（{MAX_INPUT_FILE_BYTES // (1024 ** 2):,} MB）"
            f"を超えています。\n{path}"
        )

    # PDF仕様ではヘッダーは先頭1,024バイト以内に置くことが推奨される。
    try:
        with path.open("rb") as input_file:
            header = input_file.read(1024)
    except OSError as exc:
        raise PdfMergeError(f"{label}を読み取れません。\n{path}") from exc
    if b"%PDF-" not in header:
        raise PdfMergeError(
            f"{label}はPDFヘッダーを確認できません。"
            "拡張子だけを .pdf に変更したファイルは使用できません。\n"
            f"{path}"
        )


def validate_output_path(path: Path) -> None:
    """出力ファイル名と保存先フォルダーを確認する。"""
    if path.suffix.lower() != ".pdf":
        raise PdfMergeError("出力ファイルの拡張子は .pdf にしてください。")
    if not path.parent.exists() or not path.parent.is_dir():
        raise PdfMergeError(f"出力先フォルダーが存在しません。\n{path.parent}")
    if path.exists() and not path.is_file():
        raise PdfMergeError(f"出力先は通常のファイルではありません。\n{path}")


def validate_pdf_document(document: "fitz.Document", label: str) -> None:
    """文書形式・暗号化状態・ページ数を確認する。"""
    if not document.is_pdf:
        raise PdfMergeError(f"{label}の実体はPDF形式ではありません。")
    if document.needs_pass:
        raise PdfMergeError(f"{label}はパスワードで保護されています。")
    if document.page_count < 1:
        raise PdfMergeError(f"{label}にページがありません。")
    if document.page_count > MAX_PDF_PAGES:
        raise PdfMergeError(
            f"{label}のページ数が安全上の上限（{MAX_PDF_PAGES:,}ページ）"
            f"を超えています。\nページ数: {document.page_count:,}"
        )


def get_unrotated_page_rect(page: "fitz.Page") -> "fitz.Rect":
    """挿入・clipで使用する、回転前の可視ページ領域を返す。"""
    rect = fitz.Rect(page.rect) * page.derotation_matrix
    if (
        rect.is_empty
        or rect.is_infinite
        or rect.width > MAX_PAGE_DIMENSION_POINTS
        or rect.height > MAX_PAGE_DIMENSION_POINTS
    ):
        raise PdfMergeError(
            "PDFページの寸法が不正、または安全上の上限を超えています。\n"
            f"ページサイズ: {tuple(rect)}"
        )
    return rect


def validate_rectangle(
    rect: "fitz.Rect",
    page_rect: "fitz.Rect",
    label: str,
    page_number: int,
) -> None:
    """座標が有限の有効な矩形で、対象ページ内に収まることを確認する。"""
    values = (rect.x0, rect.y0, rect.x1, rect.y1)
    if not all(math.isfinite(value) for value in values):
        raise PdfMergeError(f"{label}には有限の数値を入力してください。")
    if rect.is_empty or rect.is_infinite or rect.x0 >= rect.x1 or rect.y0 >= rect.y1:
        raise PdfMergeError(
            f"{label}が不正です。左上は右下より小さい値にしてください。\n"
            f"指定: {tuple(rect)}"
        )

    tolerance = 0.001
    is_inside = (
        rect.x0 >= page_rect.x0 - tolerance
        and rect.y0 >= page_rect.y0 - tolerance
        and rect.x1 <= page_rect.x1 + tolerance
        and rect.y1 <= page_rect.y1 + tolerance
    )
    if not is_inside:
        raise PdfMergeError(
            f"{label}が{page_number}ページ目の範囲外です。\n"
            f"指定: {tuple(rect)}\nページ: {tuple(page_rect)}"
        )


def parse_invoice_pages(page_specification: str, page_count: int) -> list[int]:
    """請求書の「すべて」「1,3-5」を0始まりページ番号へ変換する。"""
    if len(page_specification) > MAX_PAGE_SPECIFICATION_LENGTH:
        raise PdfMergeError("請求書ページの指定が長すぎます。")

    normalized = (
        page_specification.strip()
        .lower()
        .replace("，", ",")
        .replace("－", "-")
        .replace("ー", "-")
    )
    if normalized in {"", "all", "すべて", "全て", "全ページ"}:
        return list(range(page_count))

    selected_pages: set[int] = set()
    try:
        for token in normalized.split(","):
            token = token.strip()
            if not token:
                raise ValueError
            if "-" in token:
                parts = token.split("-")
                if len(parts) != 2:
                    raise ValueError
                start, end = (int(value.strip()) for value in parts)
                if start > end:
                    raise ValueError
                selected_pages.update(range(start, end + 1))
            else:
                selected_pages.add(int(token))
    except ValueError as exc:
        raise PdfMergeError(
            "請求書ページは『すべて』または『1,3-5』の形式で入力してください。"
        ) from exc

    invalid = sorted(page for page in selected_pages if not 1 <= page <= page_count)
    if invalid:
        raise PdfMergeError(
            f"請求書PDFは{page_count}ページです。範囲外の指定: {invalid}"
        )
    if not selected_pages:
        raise PdfMergeError("処理する請求書ページが指定されていません。")
    return [page - 1 for page in sorted(selected_pages)]


def create_envelope_invoice_pdf(
    invoice_path: Path,
    envelope_template_path: Path,
    output_path: Path,
    address_clip_coords: tuple[float, float, float, float],
    envelope_destination_coords: tuple[float, float, float, float],
    invoice_page_specification: str = "すべて",
    progress_callback: ProgressCallback | None = None,
) -> tuple[int, int]:
    """封筒宛名ページと請求書ページを交互に並べたPDFを作成する。

    請求書PDFは「1ページ＝1ユーザー」を前提とする。宛名は全請求書ページの
    同じ座標から抽出し、封筒テンプレートPDFの1ページ目へ配置する。

    戻り値は ``(処理したユーザー数, 出力ページ数)``。
    """
    if fitz is None:
        raise PdfMergeError(
            "PyMuPDFがインストールされていません。\n"
            "python -m pip install PyMuPDF を実行してください。"
        ) from FITZ_IMPORT_ERROR

    ensure_distinct_paths(invoice_path, envelope_template_path, output_path)
    validate_input_file(invoice_path, "請求書PDF")
    validate_input_file(envelope_template_path, "封筒テンプレートPDF")
    validate_output_path(output_path)

    address_clip = fitz.Rect(address_clip_coords)
    envelope_destination = fitz.Rect(envelope_destination_coords)
    temporary_output = output_path.with_name(
        f".{output_path.stem}_{uuid.uuid4().hex}.tmp.pdf"
    )
    expected_output_page_count: int | None = None

    def notify(message: str) -> None:
        # PyInstallerの--windowedでは標準出力がNoneになるため、存在時だけ表示する。
        if sys.stdout is not None:
            print(message)
        if progress_callback is not None:
            progress_callback(message)

    try:
        with (
            fitz.open(str(invoice_path)) as invoice_document,
            fitz.open(str(envelope_template_path)) as envelope_document,
            fitz.open() as output_document,
        ):
            validate_pdf_document(invoice_document, "請求書PDF")
            validate_pdf_document(envelope_document, "封筒テンプレートPDF")

            invoice_page_indices = parse_invoice_pages(
                invoice_page_specification,
                invoice_document.page_count,
            )
            expected_output_page_count = len(invoice_page_indices) * 2
            if expected_output_page_count > MAX_PDF_PAGES:
                raise PdfMergeError(
                    "出力PDFのページ数が安全上の上限"
                    f"（{MAX_PDF_PAGES:,}ページ）を超えます。\n"
                    f"処理ユーザー数: {len(invoice_page_indices):,}\n"
                    f"出力予定ページ数: {expected_output_page_count:,}"
                )

            envelope_template_page = envelope_document.load_page(0)
            validate_rectangle(
                envelope_destination,
                get_unrotated_page_rect(envelope_template_page),
                "封筒上の宛名配置範囲",
                1,
            )

            # 途中まで出力してから座標エラーにならないよう、全請求書ページを先に検証する。
            rotated_pages: list[tuple[int, int]] = []
            for invoice_page_index in invoice_page_indices:
                invoice_page = invoice_document.load_page(invoice_page_index)
                validate_rectangle(
                    address_clip,
                    get_unrotated_page_rect(invoice_page),
                    "宛名の切り取り範囲",
                    invoice_page_index + 1,
                )
                if invoice_page.rotation:
                    rotated_pages.append(
                        (invoice_page_index + 1, invoice_page.rotation)
                    )

            if envelope_document.page_count > 1:
                notify(
                    "注意: 封筒テンプレートPDFは複数ページですが、"
                    "1ページ目だけを使用します。"
                )
            if rotated_pages:
                rotation_summary = ", ".join(
                    f"{page_number}ページ={rotation}°"
                    for page_number, rotation in rotated_pages[:10]
                )
                if len(rotated_pages) > 10:
                    rotation_summary += f" ほか{len(rotated_pages) - 10}ページ"
                notify(f"注意: 回転情報のある請求書ページ: {rotation_summary}")

            total_users = len(invoice_page_indices)
            for sequence, invoice_page_index in enumerate(
                invoice_page_indices,
                start=1,
            ):
                is_last_user = sequence == total_users

                # 封筒テンプレートの1ページ目を追加する。
                # final=0 は、同じ元PDFを繰り返し挿入するときの重複を抑える。
                output_document.insert_pdf(
                    envelope_document,
                    from_page=0,
                    to_page=0,
                    final=1 if is_last_user else 0,
                )
                envelope_output_page = output_document.load_page(
                    output_document.page_count - 1
                )

                # このユーザーの請求書ページから宛名を抽出して封筒へ配置する。
                envelope_output_page.show_pdf_page(
                    envelope_destination,
                    invoice_document,
                    pno=invoice_page_index,
                    clip=address_clip,
                    keep_proportion=False,
                    overlay=True,
                    rotate=0,
                )

                # 元の請求書ページを、その封筒ページの直後へ追加する。
                output_document.insert_pdf(
                    invoice_document,
                    from_page=invoice_page_index,
                    to_page=invoice_page_index,
                    final=1 if is_last_user else 0,
                )
                notify(
                    f"処理中: {sequence}/{total_users} "
                    f"（請求書{invoice_page_index + 1}ページ目）"
                )

            if output_document.page_count != expected_output_page_count:
                raise PdfMergeError(
                    "出力直前のページ数が想定と一致しません。\n"
                    f"想定: {expected_output_page_count:,}ページ\n"
                    f"実際: {output_document.page_count:,}ページ"
                )

            output_document.save(
                str(temporary_output),
                garbage=4,
                deflate=True,
                clean=True,
            )

        if not temporary_output.exists() or temporary_output.stat().st_size == 0:
            raise PdfMergeError("出力PDFの一時保存に失敗しました。")

        # 保存直後に再オープンし、破損やページ欠落がないことを確認してから置換する。
        with fitz.open(str(temporary_output)) as verification_document:
            validate_pdf_document(verification_document, "出力PDF")
            if verification_document.page_count != expected_output_page_count:
                raise PdfMergeError(
                    "出力PDFのページ数が想定と一致しません。\n"
                    f"想定: {expected_output_page_count:,}ページ\n"
                    f"実際: {verification_document.page_count:,}ページ"
                )
        os.replace(temporary_output, output_path)
        return len(invoice_page_indices), expected_output_page_count
    finally:
        try:
            if temporary_output.exists():
                temporary_output.unlink()
        except OSError:
            pass


def create_clip_preview_png(
    invoice_path: Path,
    sample_page_number: int,
    address_clip_coords: tuple[float, float, float, float],
) -> tuple[bytes, str]:
    """見本ページの宛名範囲をPNG化し、説明とともに返す。"""
    if fitz is None:
        raise PdfMergeError("PyMuPDFがインストールされていません。")
    validate_input_file(invoice_path, "請求書PDF")
    if sample_page_number < 1:
        raise PdfMergeError("見本ページには1以上の整数を入力してください。")

    with fitz.open(str(invoice_path)) as document:
        validate_pdf_document(document, "請求書PDF")
        if sample_page_number > document.page_count:
            raise PdfMergeError(
                f"請求書PDFは{document.page_count}ページです。\n"
                f"指定ページ: {sample_page_number}"
            )
        page = document.load_page(sample_page_number - 1)
        clip = fitz.Rect(address_clip_coords)
        validate_rectangle(
            clip,
            get_unrotated_page_rect(page),
            "宛名の切り取り範囲",
            sample_page_number,
        )

        base_scale = 3.0
        max_width = 900.0
        max_height = 520.0
        scale = min(
            base_scale,
            max_width / max(clip.width, 1.0),
            max_height / max(clip.height, 1.0),
        )
        scale = max(scale, 0.2)
        scale = min(
            scale,
            math.sqrt(
                MAX_PREVIEW_PIXELS / max(clip.width * clip.height, 1.0)
            ),
        )
        pixmap = page.get_pixmap(
            matrix=fitz.Matrix(scale, scale),
            clip=clip,
            alpha=False,
            annots=False,
        )
        description = (
            f"見本ページ {sample_page_number}/{document.page_count}  "
            f"ページサイズ {page.rect.width:.1f} × {page.rect.height:.1f} pt  "
            f"回転 {page.rotation}°"
        )
        return pixmap.tobytes("png"), description


def render_page_preview_png(
    path: Path,
    page_number: int,
) -> tuple[bytes, int, int, "fitz.Rect", "fitz.Matrix", "fitz.Matrix", int, int]:
    """範囲選択画面用にページ全体を安全な大きさでPNG化する。

    戻り値には、画面座標とPDF座標を相互変換するためのページ矩形と
    回転／回転解除行列も含める。PDF自体は変更しない。
    """
    if fitz is None:
        raise PdfMergeError("PyMuPDFがインストールされていません。")
    validate_input_file(path, "PDF")
    if page_number < 1:
        raise PdfMergeError("ページ番号には1以上の整数を指定してください。")

    with fitz.open(str(path)) as document:
        validate_pdf_document(document, "PDF")
        if page_number > document.page_count:
            raise PdfMergeError(
                f"PDFは{document.page_count}ページです。\n指定ページ: {page_number}"
            )
        page = document.load_page(page_number - 1)
        display_rect = fitz.Rect(page.rect)
        get_unrotated_page_rect(page)  # ページ寸法の安全性を確認する。

        scale = min(
            PREVIEW_MAX_WIDTH / max(display_rect.width, 1.0),
            PREVIEW_MAX_HEIGHT / max(display_rect.height, 1.0),
            2.0,
        )
        scale = max(scale, 0.01)
        estimated_pixels = display_rect.width * display_rect.height * scale * scale
        if estimated_pixels > MAX_PREVIEW_PIXELS:
            scale *= math.sqrt(MAX_PREVIEW_PIXELS / estimated_pixels)

        pixmap = page.get_pixmap(
            matrix=fitz.Matrix(scale, scale),
            alpha=False,
            annots=False,
        )
        if pixmap.width * pixmap.height > MAX_PREVIEW_PIXELS:
            raise PdfMergeError("プレビュー画像が安全上の画素数上限を超えています。")

        return (
            pixmap.tobytes("png"),
            pixmap.width,
            pixmap.height,
            display_rect,
            fitz.Matrix(page.rotation_matrix),
            fitz.Matrix(page.derotation_matrix),
            document.page_count,
            page.rotation,
        )


class PdfRegionSelector:
    """PDFページを表示し、マウスドラッグで矩形座標を選択する画面。"""

    def __init__(
        self,
        parent: tk.Tk,
        pdf_path: Path,
        initial_page_number: int,
        initial_coordinates: tuple[float, float, float, float],
        title: str,
        instruction: str,
        outline_color: str,
    ) -> None:
        self.parent = parent
        self.pdf_path = pdf_path
        self.initial_page_number = initial_page_number
        self.initial_coordinates = initial_coordinates
        self.outline_color = outline_color
        self.result: tuple[tuple[float, float, float, float], int] | None = None

        self.page_number = initial_page_number
        self.page_count = 0
        self.page_rotation = 0
        self.image_width = 0
        self.image_height = 0
        self.display_page_rect: "fitz.Rect | None" = None
        self.rotation_matrix: "fitz.Matrix | None" = None
        self.derotation_matrix: "fitz.Matrix | None" = None
        self.selected_pdf_rect: "fitz.Rect | None" = None
        self.drag_start: tuple[float, float] | None = None
        self.selection_item: int | None = None
        self.photo: tk.PhotoImage | None = None

        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.transient(parent)
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        window_width = min(1_080, max(760, screen_width - 100))
        window_height = min(820, max(600, screen_height - 120))
        self.window.geometry(f"{window_width}x{window_height}")
        self.window.minsize(720, 560)
        self.window.protocol("WM_DELETE_WINDOW", self._cancel)
        self.window.bind("<Escape>", lambda _event: self._cancel())

        outer = ttk.Frame(self.window, padding=12)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(2, weight=1)

        ttk.Label(
            outer,
            text=instruction,
            wraplength=1_000,
            justify="left",
        ).grid(row=0, column=0, sticky="ew", pady=(0, 8))

        toolbar = ttk.Frame(outer)
        toolbar.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        ttk.Button(toolbar, text="◀ 前ページ", command=self._previous_page).pack(
            side="left"
        )
        ttk.Label(toolbar, text="ページ").pack(side="left", padx=(12, 4))
        self.page_var = tk.StringVar(value=str(initial_page_number))
        page_entry = ttk.Entry(toolbar, textvariable=self.page_var, width=7)
        page_entry.pack(side="left")
        page_entry.bind("<Return>", lambda _event: self._go_to_entered_page())
        ttk.Button(toolbar, text="表示", command=self._go_to_entered_page).pack(
            side="left", padx=(4, 0)
        )
        ttk.Button(toolbar, text="次ページ ▶", command=self._next_page).pack(
            side="left", padx=(8, 0)
        )
        self.page_info_var = tk.StringVar(value="")
        ttk.Label(toolbar, textvariable=self.page_info_var).pack(
            side="left", padx=(14, 0)
        )
        ttk.Button(toolbar, text="選択解除", command=self._clear_selection).pack(
            side="right"
        )

        canvas_frame = ttk.Frame(outer)
        canvas_frame.grid(row=2, column=0, sticky="nsew")
        canvas_frame.columnconfigure(0, weight=1)
        canvas_frame.rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(
            canvas_frame,
            background="#404040",
            cursor="crosshair",
            highlightthickness=0,
        )
        horizontal_scrollbar = ttk.Scrollbar(
            canvas_frame, orient="horizontal", command=self.canvas.xview
        )
        vertical_scrollbar = ttk.Scrollbar(
            canvas_frame, orient="vertical", command=self.canvas.yview
        )
        self.canvas.configure(
            xscrollcommand=horizontal_scrollbar.set,
            yscrollcommand=vertical_scrollbar.set,
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vertical_scrollbar.grid(row=0, column=1, sticky="ns")
        horizontal_scrollbar.grid(row=1, column=0, sticky="ew")
        self.canvas.bind("<ButtonPress-1>", self._start_selection)
        self.canvas.bind("<B1-Motion>", self._update_selection)
        self.canvas.bind("<ButtonRelease-1>", self._finish_selection)

        bottom = ttk.Frame(outer)
        bottom.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        self.selection_info_var = tk.StringVar(
            value="PDF上で左上から右下へドラッグしてください。"
        )
        ttk.Label(
            bottom,
            textvariable=self.selection_info_var,
            wraplength=760,
            justify="left",
        ).pack(side="left", fill="x", expand=True)
        ttk.Button(bottom, text="キャンセル", command=self._cancel).pack(
            side="right", padx=(8, 0)
        )
        self.accept_button = ttk.Button(
            bottom,
            text="この範囲を使用",
            command=self._accept,
            state="disabled",
        )
        self.accept_button.pack(side="right")

        try:
            self._render_page(use_initial_selection=True)
        except Exception:
            self.window.destroy()
            raise

        self.window.grab_set()
        self.window.focus_set()

    def show(self) -> tuple[tuple[float, float, float, float], int] | None:
        """選択画面をモーダル表示し、確定された座標とページ番号を返す。"""
        self.parent.wait_window(self.window)
        return self.result

    def _render_page(self, use_initial_selection: bool = False) -> None:
        (
            png_bytes,
            self.image_width,
            self.image_height,
            self.display_page_rect,
            self.rotation_matrix,
            self.derotation_matrix,
            self.page_count,
            self.page_rotation,
        ) = render_page_preview_png(self.pdf_path, self.page_number)

        image_data = base64.b64encode(png_bytes).decode("ascii")
        self.photo = tk.PhotoImage(data=image_data, format="png")
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, image=self.photo, anchor="nw", tags="page")
        self.canvas.configure(
            scrollregion=(0, 0, self.image_width, self.image_height)
        )
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)
        self.page_var.set(str(self.page_number))
        self.page_info_var.set(
            f"/ {self.page_count}　表示回転: {self.page_rotation}°"
        )
        self.drag_start = None
        self.selection_item = None
        self.selected_pdf_rect = None
        self.accept_button.configure(state="disabled")
        self.selection_info_var.set(
            "PDF上で左上から右下へドラッグしてください。"
        )

        if use_initial_selection:
            initial_rect = fitz.Rect(self.initial_coordinates)
            try:
                unrotated_page_rect = (
                    fitz.Rect(self.display_page_rect) * self.derotation_matrix
                )
                validate_rectangle(
                    initial_rect,
                    unrotated_page_rect,
                    "現在の座標",
                    self.page_number,
                )
            except PdfMergeError:
                return
            self.selected_pdf_rect = initial_rect
            self._draw_pdf_selection(initial_rect)
            self._update_selection_information(initial_rect)
            self.accept_button.configure(state="normal")

    def _event_canvas_point(self, event: tk.Event) -> tuple[float, float]:
        """イベント座標を画像内へ収めたCanvas座標に変換する。"""
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        return (
            min(max(x, 0.0), float(self.image_width)),
            min(max(y, 0.0), float(self.image_height)),
        )

    def _start_selection(self, event: tk.Event) -> None:
        if self.photo is None:
            return
        self.drag_start = self._event_canvas_point(event)
        if self.selection_item is not None:
            self.canvas.delete(self.selection_item)
        x, y = self.drag_start
        self.selection_item = self.canvas.create_rectangle(
            x,
            y,
            x,
            y,
            outline=self.outline_color,
            width=3,
            dash=(7, 3),
            tags="selection",
        )
        self.selected_pdf_rect = None
        self.accept_button.configure(state="disabled")

    def _update_selection(self, event: tk.Event) -> None:
        if self.drag_start is None or self.selection_item is None:
            return
        current_x, current_y = self._event_canvas_point(event)
        start_x, start_y = self.drag_start
        self.canvas.coords(
            self.selection_item,
            min(start_x, current_x),
            min(start_y, current_y),
            max(start_x, current_x),
            max(start_y, current_y),
        )

    def _finish_selection(self, event: tk.Event) -> None:
        if self.drag_start is None:
            return
        end_x, end_y = self._event_canvas_point(event)
        start_x, start_y = self.drag_start
        self.drag_start = None
        if abs(end_x - start_x) < 3 or abs(end_y - start_y) < 3:
            self._clear_selection()
            self.selection_info_var.set(
                "選択範囲が小さすぎます。範囲をドラッグして指定してください。"
            )
            return

        selected_rect = self._canvas_rect_to_pdf_rect(
            min(start_x, end_x),
            min(start_y, end_y),
            max(start_x, end_x),
            max(start_y, end_y),
        )
        self.selected_pdf_rect = selected_rect
        self._draw_pdf_selection(selected_rect)
        self._update_selection_information(selected_rect)
        self.accept_button.configure(state="normal")

    def _canvas_rect_to_pdf_rect(
        self, x0: float, y0: float, x1: float, y1: float
    ) -> "fitz.Rect":
        """表示中のCanvas矩形を、回転前のPDF座標へ変換する。"""
        if self.display_page_rect is None or self.derotation_matrix is None:
            raise PdfMergeError("プレビューの座標情報を取得できません。")
        scale_x = self.display_page_rect.width / self.image_width
        scale_y = self.display_page_rect.height / self.image_height
        displayed_pdf_rect = fitz.Rect(
            self.display_page_rect.x0 + x0 * scale_x,
            self.display_page_rect.y0 + y0 * scale_y,
            self.display_page_rect.x0 + x1 * scale_x,
            self.display_page_rect.y0 + y1 * scale_y,
        )
        unrotated_rect = displayed_pdf_rect * self.derotation_matrix
        return fitz.Rect(
            min(unrotated_rect.x0, unrotated_rect.x1),
            min(unrotated_rect.y0, unrotated_rect.y1),
            max(unrotated_rect.x0, unrotated_rect.x1),
            max(unrotated_rect.y0, unrotated_rect.y1),
        )

    def _draw_pdf_selection(self, rect: "fitz.Rect") -> None:
        """回転前のPDF矩形を画面座標へ変換して選択枠を描画する。"""
        if self.display_page_rect is None or self.rotation_matrix is None:
            return
        displayed_rect = fitz.Rect(rect) * self.rotation_matrix
        scale_x = self.image_width / self.display_page_rect.width
        scale_y = self.image_height / self.display_page_rect.height
        canvas_rect = (
            (displayed_rect.x0 - self.display_page_rect.x0) * scale_x,
            (displayed_rect.y0 - self.display_page_rect.y0) * scale_y,
            (displayed_rect.x1 - self.display_page_rect.x0) * scale_x,
            (displayed_rect.y1 - self.display_page_rect.y0) * scale_y,
        )
        if self.selection_item is None:
            self.selection_item = self.canvas.create_rectangle(
                *canvas_rect,
                outline=self.outline_color,
                width=3,
                dash=(7, 3),
                tags="selection",
            )
        else:
            self.canvas.coords(self.selection_item, *canvas_rect)

    def _update_selection_information(self, rect: "fitz.Rect") -> None:
        self.selection_info_var.set(
            "選択座標: "
            f"({rect.x0:.2f}, {rect.y0:.2f}) - "
            f"({rect.x1:.2f}, {rect.y1:.2f}) pt"
        )

    def _clear_selection(self) -> None:
        if self.selection_item is not None:
            self.canvas.delete(self.selection_item)
        self.selection_item = None
        self.selected_pdf_rect = None
        self.drag_start = None
        self.accept_button.configure(state="disabled")
        self.selection_info_var.set(
            "PDF上で左上から右下へドラッグしてください。"
        )

    def _parse_entered_page(self) -> int:
        try:
            page_number = int(self.page_var.get().strip())
        except ValueError as exc:
            raise PdfMergeError("ページ番号には整数を入力してください。") from exc
        if not 1 <= page_number <= self.page_count:
            raise PdfMergeError(
                f"ページ番号は1から{self.page_count}の範囲で指定してください。"
            )
        return page_number

    def _go_to_entered_page(self) -> None:
        try:
            self._change_page(self._parse_entered_page())
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc), parent=self.window)

    def _previous_page(self) -> None:
        if self.page_number <= 1:
            return
        self._change_page(self.page_number - 1)

    def _next_page(self) -> None:
        if self.page_number >= self.page_count:
            return
        self._change_page(self.page_number + 1)

    def _change_page(self, new_page_number: int) -> None:
        """表示失敗時に元ページへ戻せるよう、安全にページを切り替える。"""
        previous_page_number = self.page_number
        self.page_number = new_page_number
        try:
            self._render_page()
        except Exception as exc:
            self.page_number = previous_page_number
            self.page_var.set(str(previous_page_number))
            messagebox.showerror(
                APP_TITLE,
                f"ページを表示できませんでした。\n{type(exc).__name__}: {exc}",
                parent=self.window,
            )

    def _accept(self) -> None:
        if self.selected_pdf_rect is None:
            messagebox.showwarning(
                APP_TITLE,
                "PDF上で範囲を選択してください。",
                parent=self.window,
            )
            return
        rect = self.selected_pdf_rect
        self.result = (
            (rect.x0, rect.y0, rect.x1, rect.y1),
            self.page_number,
        )
        self.window.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.window.destroy()


class PdfMergeApp:
    """請求書と封筒の設定、プレビュー、PDF作成を提供するGUI。"""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("900x720")
        self.root.minsize(800, 660)

        application_directory = get_application_directory()
        default_invoice = application_directory / "invoice.pdf"
        default_envelope = application_directory / "envelope_template.pdf"
        default_output = application_directory / "invoice_envelope_output.pdf"

        self.invoice_path_var = tk.StringVar(
            value=str(default_invoice) if default_invoice.exists() else ""
        )
        self.envelope_path_var = tk.StringVar(
            value=str(default_envelope) if default_envelope.exists() else ""
        )
        self.output_path_var = tk.StringVar(value=str(default_output))
        self.sample_page_var = tk.StringVar(value="1")
        self.invoice_pages_var = tk.StringVar(value="すべて")
        self.status_var = tk.StringVar(
            value="請求書PDF、封筒テンプレート、宛名範囲を指定してください。"
        )
        self.last_output_path: Path | None = None
        self.address_coord_vars = [
            tk.StringVar(value=f"{value:g}") for value in DEFAULT_ADDRESS_COORDS
        ]
        self.envelope_coord_vars = [
            tk.StringVar(value=f"{value:g}")
            for value in DEFAULT_ENVELOPE_DESTINATION_COORDS
        ]
        self._build_ui()

    def _build_ui(self) -> None:
        """画面部品を作成する。"""
        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)

        ttk.Label(
            outer,
            text=APP_TITLE,
            font=("Yu Gothic UI", 16, "bold"),
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))
        ttk.Label(
            outer,
            text=(
                "請求書の各ページから宛名を抽出し、"
                "封筒宛名ページと請求書ページを交互に作成します。"
                "前提: 請求書PDFの1ページが1ユーザー分です。"
            ),
            wraplength=820,
        ).grid(row=1, column=0, sticky="w", pady=(0, 12))

        files_frame = ttk.LabelFrame(outer, text="1. ファイル", padding=10)
        files_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        files_frame.columnconfigure(1, weight=1)
        self._add_file_row(
            files_frame,
            0,
            "請求書PDF",
            self.invoice_path_var,
            self._select_invoice_pdf,
        )
        self._add_file_row(
            files_frame,
            1,
            "封筒テンプレート",
            self.envelope_path_var,
            self._select_envelope_template_pdf,
        )
        self._add_file_row(files_frame, 2, "出力PDF", self.output_path_var, self._select_output_pdf)

        source_frame = ttk.LabelFrame(
            outer,
            text="2. 宛名の切り取り設定（請求書）",
            padding=10,
        )
        source_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(source_frame, text="見本ページ").grid(
            row=0, column=0, sticky="w", padx=(0, 6), pady=4
        )
        ttk.Entry(source_frame, textvariable=self.sample_page_var, width=8).grid(
            row=0, column=1, sticky="w", pady=4
        )
        ttk.Label(source_frame, text="（宛名範囲を決めるためのページ）").grid(
            row=0, column=2, sticky="w", padx=(4, 14), pady=4
        )
        self._add_coordinate_row(
            source_frame,
            1,
            "宛名座標",
            self.address_coord_vars,
        )
        source_button_frame = ttk.Frame(source_frame)
        source_button_frame.grid(
            row=2, column=0, columnspan=6, sticky="w", pady=(8, 2)
        )
        ttk.Button(
            source_button_frame,
            text="請求書上で宛名範囲を指定",
            command=self._select_address_region_on_pdf,
        ).pack(side="left")
        ttk.Button(
            source_button_frame,
            text="宛名プレビュー",
            command=self._show_address_preview,
        ).pack(side="left", padx=(8, 0))

        destination_frame = ttk.LabelFrame(
            outer, text="3. 宛名の配置設定（封筒）", padding=10
        )
        destination_frame.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        self._add_coordinate_row(
            destination_frame,
            0,
            "封筒上の座標",
            self.envelope_coord_vars,
        )
        ttk.Label(destination_frame, text="処理する請求書ページ").grid(
            row=1, column=0, sticky="w", padx=(0, 6), pady=4
        )
        ttk.Entry(
            destination_frame, textvariable=self.invoice_pages_var, width=20
        ).grid(row=1, column=1, columnspan=2, sticky="w", pady=4)
        ttk.Label(destination_frame, text="例: すべて / 1 / 1,3-5").grid(
            row=1, column=3, columnspan=3, sticky="w", padx=(8, 0), pady=4
        )
        ttk.Button(
            destination_frame,
            text="封筒上で宛名の配置範囲を指定",
            command=self._select_envelope_region_on_pdf,
        ).grid(row=2, column=0, columnspan=6, sticky="w", pady=(8, 2))

        ttk.Label(
            outer,
            text=(
                "座標単位は point、原点はページ左上です。"
                "1 mm ≒ 2.83465 point。マウス指定後も数値欄で微調整できます。"
            ),
            foreground="#555555",
            wraplength=760,
        ).grid(row=5, column=0, sticky="w", pady=(0, 10))

        action_frame = ttk.Frame(outer)
        action_frame.grid(row=6, column=0, sticky="ew", pady=(2, 10))
        self.run_button = ttk.Button(
            action_frame, text="宛名・請求書PDFを作成", command=self._start_merge
        )
        self.run_button.pack(side="left")
        self.open_button = ttk.Button(
            action_frame,
            text="出力PDFを開く",
            command=self._open_output_pdf,
            state="disabled",
        )
        self.open_button.pack(side="left", padx=(8, 0))
        ttk.Button(action_frame, text="閉じる", command=self.root.destroy).pack(side="right")

        status_frame = ttk.LabelFrame(outer, text="状況", padding=10)
        status_frame.grid(row=7, column=0, sticky="nsew")
        outer.rowconfigure(7, weight=1)
        ttk.Label(
            status_frame,
            textvariable=self.status_var,
            wraplength=750,
            justify="left",
        ).pack(fill="both", expand=True, anchor="nw")

    @staticmethod
    def _add_file_row(
        parent: ttk.LabelFrame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command: Callable[[], None],
    ) -> None:
        ttk.Label(parent, text=label, width=12).grid(
            row=row, column=0, sticky="w", padx=(0, 6), pady=4
        )
        ttk.Entry(parent, textvariable=variable).grid(
            row=row, column=1, sticky="ew", pady=4
        )
        ttk.Button(parent, text="参照...", command=command).grid(
            row=row, column=2, sticky="e", padx=(8, 0), pady=4
        )

    @staticmethod
    def _add_coordinate_row(
        parent: ttk.LabelFrame,
        row: int,
        label: str,
        variables: list[tk.StringVar],
    ) -> None:
        labels = ("左上X", "左上Y", "右下X", "右下Y")
        ttk.Label(parent, text=label).grid(
            row=row, column=0, sticky="w", padx=(0, 6), pady=4
        )
        for index, (coordinate_label, variable) in enumerate(
            zip(labels, variables, strict=True), start=1
        ):
            frame = ttk.Frame(parent)
            frame.grid(row=row, column=index, sticky="w", padx=(2, 6), pady=4)
            ttk.Label(frame, text=coordinate_label).pack(anchor="w")
            ttk.Entry(frame, textvariable=variable, width=10).pack(anchor="w")

    def _dialog_initial_directory(self, current_value: str) -> str:
        path = Path(current_value.strip()) if current_value.strip() else None
        if path is not None and path.parent.exists():
            return str(path.parent)
        return str(get_application_directory())

    def _select_invoice_pdf(self) -> None:
        selected = filedialog.askopenfilename(
            title="請求書PDFを選択",
            initialdir=self._dialog_initial_directory(self.invoice_path_var.get()),
            filetypes=PDF_FILE_TYPES,
        )
        if selected:
            self.invoice_path_var.set(selected)
            invoice_path = Path(selected)
            current_output = self.output_path_var.get().strip()
            default_name = get_application_directory() / "invoice_envelope_output.pdf"
            if not current_output or Path(current_output) == default_name:
                self.output_path_var.set(
                    str(invoice_path.with_name(f"{invoice_path.stem}_envelopes.pdf"))
                )
            self.status_var.set(
                "請求書PDFを選択しました。見本ページで宛名範囲を指定してください。"
            )

    def _select_envelope_template_pdf(self) -> None:
        selected = filedialog.askopenfilename(
            title="封筒テンプレートPDFを選択",
            initialdir=self._dialog_initial_directory(self.envelope_path_var.get()),
            filetypes=PDF_FILE_TYPES,
        )
        if selected:
            self.envelope_path_var.set(selected)
            self.status_var.set(
                "封筒テンプレートPDFを選択しました。宛名の配置範囲を指定してください。"
            )

    def _select_output_pdf(self) -> None:
        selected = filedialog.asksaveasfilename(
            title="出力PDFを指定",
            initialdir=self._dialog_initial_directory(self.output_path_var.get()),
            initialfile=(
                Path(self.output_path_var.get()).name
                if self.output_path_var.get().strip()
                else "invoice_envelope_output.pdf"
            ),
            defaultextension=".pdf",
            filetypes=PDF_FILE_TYPES,
        )
        if selected:
            self.output_path_var.set(selected)
            self.status_var.set("出力先を指定しました。")

    @staticmethod
    def _set_coordinate_variables(
        variables: list[tk.StringVar],
        coordinates: tuple[float, float, float, float],
    ) -> None:
        """マウス選択結果を読みやすい小数表記で座標欄へ設定する。"""
        for variable, value in zip(variables, coordinates, strict=True):
            variable.set(f"{value:.2f}".rstrip("0").rstrip("."))

    def _select_address_region_on_pdf(self) -> None:
        """請求書を表示し、宛名範囲と見本ページをマウスで選択する。"""
        try:
            invoice_text = self.invoice_path_var.get().strip()
            if not invoice_text:
                raise PdfMergeError("請求書PDFを選択してください。")
            page_number = self._parse_positive_integer(
                self.sample_page_var.get(), "見本ページ"
            )
            initial_coordinates = self._parse_coordinates(
                self.address_coord_vars, "宛名座標"
            )
            selector = PdfRegionSelector(
                parent=self.root,
                pdf_path=Path(invoice_text),
                initial_page_number=page_number,
                initial_coordinates=initial_coordinates,
                title="請求書上の宛名範囲を指定",
                instruction=(
                    "請求書の宛名部分を、左上から右下へドラッグしてください。"
                    "ここで決めた座標を、処理対象の全請求書ページに適用します。"
                ),
                outline_color="#e53935",
            )
            result = selector.show()
            if result is None:
                self.status_var.set("宛名範囲の指定をキャンセルしました。")
                return
            coordinates, selected_page_number = result
            self._set_coordinate_variables(self.address_coord_vars, coordinates)
            self.sample_page_var.set(str(selected_page_number))
            self.status_var.set(
                f"請求書{selected_page_number}ページ目を見本に宛名範囲を設定しました。"
            )
        except Exception as exc:
            self._show_error(exc)

    def _select_envelope_region_on_pdf(self) -> None:
        """封筒テンプレートを表示し、宛名の配置範囲を選択する。"""
        try:
            envelope_text = self.envelope_path_var.get().strip()
            if not envelope_text:
                raise PdfMergeError("封筒テンプレートPDFを選択してください。")
            envelope_path = Path(envelope_text)
            validate_input_file(envelope_path, "封筒テンプレートPDF")
            with fitz.open(str(envelope_path)) as envelope_document:
                validate_pdf_document(envelope_document, "封筒テンプレートPDF")
            initial_coordinates = self._parse_coordinates(
                self.envelope_coord_vars, "封筒上の宛名配置座標"
            )
            selector = PdfRegionSelector(
                parent=self.root,
                pdf_path=envelope_path,
                initial_page_number=1,
                initial_coordinates=initial_coordinates,
                title="封筒上の宛名配置範囲を指定",
                instruction=(
                    "宛名を印刷したい範囲を、左上から右下へドラッグしてください。"
                    "封筒テンプレートの1ページ目を使用します。"
                ),
                outline_color="#1976d2",
            )
            result = selector.show()
            if result is None:
                self.status_var.set("封筒上の配置範囲の指定をキャンセルしました。")
                return
            coordinates, preview_page_number = result
            if preview_page_number != 1:
                raise PdfMergeError(
                    "封筒の宛名配置範囲は、テンプレートの1ページ目で"
                    "指定してください。"
                )
            self._set_coordinate_variables(
                self.envelope_coord_vars, coordinates
            )
            self.status_var.set("封筒テンプレート上の宛名配置範囲を設定しました。")
        except Exception as exc:
            self._show_error(exc)

    @staticmethod
    def _parse_positive_integer(value: str, label: str) -> int:
        try:
            number = int(value.strip())
        except ValueError as exc:
            raise PdfMergeError(f"{label}には整数を入力してください。") from exc
        if number < 1:
            raise PdfMergeError(f"{label}には1以上の整数を入力してください。")
        return number

    @staticmethod
    def _parse_coordinates(
        variables: list[tk.StringVar], label: str
    ) -> tuple[float, float, float, float]:
        try:
            values = tuple(float(variable.get().strip()) for variable in variables)
        except ValueError as exc:
            raise PdfMergeError(f"{label}には数値を入力してください。") from exc
        if len(values) != 4 or not all(math.isfinite(value) for value in values):
            raise PdfMergeError(f"{label}には有限の数値を4つ入力してください。")
        return values  # type: ignore[return-value]

    def _collect_inputs(
        self,
    ) -> tuple[
        Path,
        Path,
        Path,
        tuple[float, float, float, float],
        tuple[float, float, float, float],
        str,
    ]:
        invoice_text = self.invoice_path_var.get().strip()
        envelope_text = self.envelope_path_var.get().strip()
        output_text = self.output_path_var.get().strip()
        if not invoice_text:
            raise PdfMergeError("請求書PDFを選択してください。")
        if not envelope_text:
            raise PdfMergeError("封筒テンプレートPDFを選択してください。")
        if not output_text:
            raise PdfMergeError("出力PDFを指定してください。")
        return (
            Path(invoice_text),
            Path(envelope_text),
            Path(output_text),
            self._parse_coordinates(self.address_coord_vars, "宛名座標"),
            self._parse_coordinates(
                self.envelope_coord_vars,
                "封筒上の宛名配置座標",
            ),
            self.invoice_pages_var.get().strip(),
        )

    def _show_address_preview(self) -> None:
        try:
            invoice_text = self.invoice_path_var.get().strip()
            if not invoice_text:
                raise PdfMergeError("請求書PDFを選択してください。")
            page_number = self._parse_positive_integer(
                self.sample_page_var.get(), "見本ページ"
            )
            coordinates = self._parse_coordinates(
                self.address_coord_vars, "宛名座標"
            )
            png_bytes, description = create_clip_preview_png(
                Path(invoice_text), page_number, coordinates
            )
        except Exception as exc:
            self._show_error(exc)
            return

        preview_window = tk.Toplevel(self.root)
        preview_window.title("宛名プレビュー")
        preview_window.transient(self.root)
        preview_window.grab_set()
        ttk.Label(preview_window, text=description, padding=(12, 10)).pack(anchor="w")
        image_data = base64.b64encode(png_bytes).decode("ascii")
        photo = tk.PhotoImage(data=image_data, format="png")
        image_label = ttk.Label(preview_window, image=photo, padding=12)
        image_label.image = photo  # type: ignore[attr-defined]  # GC防止
        image_label.pack(fill="both", expand=True)
        ttk.Label(
            preview_window,
            text=(
                "この範囲と同じ座標を、処理対象の全請求書ページから"
                "宛名として抽出します。"
            ),
            padding=(12, 0, 12, 8),
        ).pack(anchor="w")
        ttk.Button(
            preview_window, text="閉じる", command=preview_window.destroy
        ).pack(pady=(0, 12))

    def _start_merge(self) -> None:
        try:
            inputs = self._collect_inputs()
            output_path = inputs[2]
            if output_path.exists() and not messagebox.askyesno(
                APP_TITLE,
                f"出力ファイルは既に存在します。上書きしますか？\n\n{output_path}",
                parent=self.root,
            ):
                self.status_var.set("処理をキャンセルしました。")
                return
        except Exception as exc:
            self._show_error(exc)
            return

        self.run_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.status_var.set("封筒宛名ページと請求書ページの作成を開始します...")
        threading.Thread(
            target=self._merge_worker, args=(inputs,), daemon=True
        ).start()

    def _merge_worker(
        self,
        inputs: tuple[
            Path,
            Path,
            Path,
            tuple[float, float, float, float],
            tuple[float, float, float, float],
            str,
        ],
    ) -> None:
        try:
            user_count, output_page_count = create_envelope_invoice_pdf(
                invoice_path=inputs[0],
                envelope_template_path=inputs[1],
                output_path=inputs[2],
                address_clip_coords=inputs[3],
                envelope_destination_coords=inputs[4],
                invoice_page_specification=inputs[5],
                progress_callback=self._post_status,
            )
        except Exception as exc:
            self.root.after(0, self._merge_failed, exc)
            return
        self.root.after(
            0,
            self._merge_succeeded,
            inputs[2],
            user_count,
            output_page_count,
        )

    def _post_status(self, message: str) -> None:
        self.root.after(0, self.status_var.set, message)

    def _merge_succeeded(
        self,
        output_path: Path,
        user_count: int,
        output_page_count: int,
    ) -> None:
        self.last_output_path = output_path
        self.run_button.configure(state="normal")
        self.open_button.configure(state="normal")
        self.status_var.set(
            f"完了: {user_count}ユーザー、{output_page_count}ページを作成しました。\n"
            f"{output_path}"
        )
        messagebox.showinfo(
            APP_TITLE,
            "宛名・請求書PDFの作成が完了しました。\n\n"
            f"処理ユーザー数: {user_count}\n"
            f"出力ページ数: {output_page_count}\n"
            f"{output_path}",
            parent=self.root,
        )

    def _merge_failed(self, exc: Exception) -> None:
        self.run_button.configure(state="normal")
        self.open_button.configure(state="disabled")
        self._show_error(exc)

    def _show_error(self, exc: Exception) -> None:
        if isinstance(exc, PdfMergeError):
            message = str(exc)
        elif isinstance(exc, PermissionError):
            message = (
                "ファイルへアクセスできません。PDFが別のアプリで開かれていないか、"
                "保存先への書き込み権限があるか確認してください。\n\n"
                f"詳細: {exc}"
            )
        else:
            message = f"予期しないエラーが発生しました。\n{type(exc).__name__}: {exc}"
        self.status_var.set(f"エラー: {message}")
        messagebox.showerror(APP_TITLE, message, parent=self.root)

    def _open_output_pdf(self) -> None:
        if self.last_output_path is None or not self.last_output_path.exists():
            messagebox.showwarning(APP_TITLE, "出力PDFが見つかりません。", parent=self.root)
            return
        try:
            if os.name != "nt":
                raise PdfMergeError("出力PDFを開く機能はWindows用です。")
            os.startfile(self.last_output_path)  # type: ignore[attr-defined]
        except Exception as exc:
            self._show_error(exc)


def main() -> int:
    """GUIを起動する。"""
    root = tk.Tk()
    if fitz is None:
        root.withdraw()
        messagebox.showerror(
            APP_TITLE,
            "PyMuPDFがインストールされていません。\n"
            "python -m pip install PyMuPDF を実行してください。\n\n"
            f"詳細: {FITZ_IMPORT_ERROR}",
            parent=root,
        )
        root.destroy()
        return 1

    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        PdfMergeApp(root)
        root.mainloop()
        return 0
    except Exception as exc:
        root.withdraw()
        messagebox.showerror(
            APP_TITLE,
            f"アプリケーションを起動できませんでした。\n{type(exc).__name__}: {exc}",
            parent=root,
        )
        root.destroy()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
