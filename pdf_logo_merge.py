#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF領域合成ツール（GUI版）

抽出元PDFの指定領域を、合成先PDFの指定位置へベクターデータのまま配置する。
Windows用EXE（PyInstaller --onefile --windowed）での利用を想定している。

主な機能:
    - 抽出元PDF、合成先PDF、出力先をファイル選択ダイアログで指定
    - 抽出元ページと切り取り座標を画面で入力
    - 貼り付け座標を画面で入力
    - 全ページまたは「1,3-5」のような指定ページへ合成
    - 切り取り範囲のプレビュー
    - 入力・座標・ページ範囲・暗号化PDF・保存先を事前検証
    - 一時保存後に出力を置き換え、保存失敗による既存PDF破損を防止

座標系:
    PyMuPDFの座標（単位: point、原点: ページ左上）を使用する。
    1 point = 1/72 inch、1 mm ≒ 2.83465 point。

重要:
    show_pdf_page() を使用するため、抽出元がベクターデータなら、
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


# PyMuPDFが未導入でも、GUIで分かりやすいエラーを表示できるようにする。
try:
    import fitz  # PyMuPDF
except ImportError as exc:  # pragma: no cover - PyMuPDF未導入時だけ通る
    fitz = None  # type: ignore[assignment]
    FITZ_IMPORT_ERROR: ImportError | None = exc
else:
    FITZ_IMPORT_ERROR = None


APP_TITLE: Final[str] = "PDF領域合成ツール"
PDF_FILE_TYPES: Final[tuple[tuple[str, str], ...]] = (
    ("PDFファイル", "*.pdf"),
    ("すべてのファイル", "*.*"),
)

# 前回の固定設定をGUIの初期値として維持する。
DEFAULT_SOURCE_COORDS: Final[tuple[float, float, float, float]] = (
    10.0,
    10.0,
    150.0,
    50.0,
)
DEFAULT_DESTINATION_COORDS: Final[tuple[float, float, float, float]] = (
    400.0,
    20.0,
    540.0,
    60.0,
)

ProgressCallback = Callable[[str], None]


class PdfMergeError(Exception):
    """画面へそのまま表示できる、利用者向けの処理エラー。"""


def get_application_directory() -> Path:
    """Python実行時はスクリプト、EXE実行時はEXEがあるフォルダーを返す。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(sys.argv[0]).resolve().parent


def ensure_distinct_paths(*paths: Path) -> None:
    """抽出元・合成先・出力先が同じファイルでないことを確認する。"""
    normalized = [os.path.normcase(str(path.resolve(strict=False))) for path in paths]
    if len(normalized) != len(set(normalized)):
        raise PdfMergeError(
            "抽出元PDF、合成先PDF、出力PDFには異なるファイルを指定してください。"
        )


def validate_input_file(path: Path, label: str) -> None:
    """入力ファイルの存在・種類・拡張子を確認する。"""
    if not path.exists():
        raise PdfMergeError(f"{label}が見つかりません。\n{path}")
    if not path.is_file():
        raise PdfMergeError(f"{label}はファイルではありません。\n{path}")
    if path.suffix.lower() != ".pdf":
        raise PdfMergeError(f"{label}にはPDFファイルを指定してください。\n{path}")


def validate_output_path(path: Path) -> None:
    """出力ファイル名と保存先フォルダーを確認する。"""
    if path.suffix.lower() != ".pdf":
        raise PdfMergeError("出力ファイルの拡張子は .pdf にしてください。")
    if not path.parent.exists() or not path.parent.is_dir():
        raise PdfMergeError(f"出力先フォルダーが存在しません。\n{path.parent}")


def validate_pdf_document(document: "fitz.Document", label: str) -> None:
    """PDFがページを持ち、パスワードなしで処理できることを確認する。"""
    if document.needs_pass:
        raise PdfMergeError(f"{label}はパスワードで保護されています。")
    if document.page_count < 1:
        raise PdfMergeError(f"{label}にページがありません。")


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


def parse_target_pages(page_specification: str, page_count: int) -> list[int]:
    """「すべて」「1,3-5」をPyMuPDF用の0始まりページ番号へ変換する。"""
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
            "対象ページは『すべて』または『1,3-5』の形式で入力してください。"
        ) from exc

    invalid = sorted(page for page in selected_pages if not 1 <= page <= page_count)
    if invalid:
        raise PdfMergeError(
            f"合成先PDFは{page_count}ページです。範囲外の指定: {invalid}"
        )
    if not selected_pages:
        raise PdfMergeError("合成するページが指定されていません。")
    return [page - 1 for page in sorted(selected_pages)]


def merge_pdf_area(
    source_path: Path,
    target_path: Path,
    output_path: Path,
    source_page_number: int,
    source_clip_coords: tuple[float, float, float, float],
    destination_coords: tuple[float, float, float, float],
    target_page_specification: str = "すべて",
    progress_callback: ProgressCallback | None = None,
) -> int:
    """指定されたPDF・座標・ページに対してベクター合成を実行する。"""
    if fitz is None:
        raise PdfMergeError(
            "PyMuPDFがインストールされていません。\n"
            "python -m pip install PyMuPDF を実行してください。"
        ) from FITZ_IMPORT_ERROR
    if source_page_number < 1:
        raise PdfMergeError("抽出元ページには1以上の整数を入力してください。")

    ensure_distinct_paths(source_path, target_path, output_path)
    validate_input_file(source_path, "抽出元PDF")
    validate_input_file(target_path, "合成先PDF")
    validate_output_path(output_path)

    source_clip = fitz.Rect(source_clip_coords)
    destination_rect = fitz.Rect(destination_coords)
    temporary_output = output_path.with_name(
        f".{output_path.stem}_{uuid.uuid4().hex}.tmp.pdf"
    )

    def notify(message: str) -> None:
        # PyInstallerの--windowedでは標準出力がNoneになるため、存在時だけ表示する。
        if sys.stdout is not None:
            print(message)
        if progress_callback is not None:
            progress_callback(message)

    try:
        with fitz.open(str(source_path)) as source_document, fitz.open(
            str(target_path)
        ) as target_document:
            validate_pdf_document(source_document, "抽出元PDF")
            validate_pdf_document(target_document, "合成先PDF")

            if source_page_number > source_document.page_count:
                raise PdfMergeError(
                    f"抽出元PDFは{source_document.page_count}ページです。\n"
                    f"指定ページ: {source_page_number}"
                )

            source_page_index = source_page_number - 1
            source_page = source_document.load_page(source_page_index)
            validate_rectangle(
                source_clip,
                source_page.rect,
                "切り取り範囲",
                source_page_number,
            )
            page_indices = parse_target_pages(
                target_page_specification,
                target_document.page_count,
            )

            # 途中まで処理してから座標エラーにならないよう、全ページを先に検証する。
            for page_index in page_indices:
                target_page = target_document.load_page(page_index)
                validate_rectangle(
                    destination_rect,
                    target_page.rect,
                    "貼り付け範囲",
                    page_index + 1,
                )

            if source_page.rotation:
                notify(
                    f"注意: 抽出元{source_page_number}ページ目には"
                    f"{source_page.rotation}度の回転情報があります。"
                )

            for sequence, page_index in enumerate(page_indices, start=1):
                target_page = target_document.load_page(page_index)
                target_page.show_pdf_page(
                    destination_rect,
                    source_document,
                    pno=source_page_index,
                    clip=source_clip,
                    keep_proportion=False,
                    overlay=True,
                    rotate=0,
                )
                notify(
                    f"処理中: {sequence}/{len(page_indices)} "
                    f"（合成先{page_index + 1}ページ目）"
                )

            target_document.save(
                str(temporary_output),
                garbage=4,
                deflate=True,
                clean=True,
            )

        if not temporary_output.exists() or temporary_output.stat().st_size == 0:
            raise PdfMergeError("出力PDFの一時保存に失敗しました。")
        os.replace(temporary_output, output_path)
        return len(page_indices)
    finally:
        try:
            if temporary_output.exists():
                temporary_output.unlink()
        except OSError:
            pass


def create_clip_preview_png(
    source_path: Path,
    source_page_number: int,
    source_clip_coords: tuple[float, float, float, float],
) -> tuple[bytes, str]:
    """切り取り範囲をPNG化し、GUIプレビュー用のバイト列と説明を返す。"""
    if fitz is None:
        raise PdfMergeError("PyMuPDFがインストールされていません。")
    validate_input_file(source_path, "抽出元PDF")
    if source_page_number < 1:
        raise PdfMergeError("抽出元ページには1以上の整数を入力してください。")

    with fitz.open(str(source_path)) as document:
        validate_pdf_document(document, "抽出元PDF")
        if source_page_number > document.page_count:
            raise PdfMergeError(
                f"抽出元PDFは{document.page_count}ページです。\n"
                f"指定ページ: {source_page_number}"
            )
        page = document.load_page(source_page_number - 1)
        clip = fitz.Rect(source_clip_coords)
        validate_rectangle(clip, page.rect, "切り取り範囲", source_page_number)

        base_scale = 3.0
        max_width = 900.0
        max_height = 520.0
        scale = min(
            base_scale,
            max_width / max(clip.width, 1.0),
            max_height / max(clip.height, 1.0),
        )
        scale = max(scale, 0.2)
        pixmap = page.get_pixmap(
            matrix=fitz.Matrix(scale, scale),
            clip=clip,
            alpha=False,
        )
        description = (
            f"ページ {source_page_number}/{document.page_count}  "
            f"ページサイズ {page.rect.width:.1f} × {page.rect.height:.1f} pt  "
            f"回転 {page.rotation}°"
        )
        return pixmap.tobytes("png"), description


class PdfMergeApp:
    """PDF合成条件の入力、プレビュー、実行を提供するTkinter GUI。"""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("820x660")
        self.root.minsize(760, 620)

        application_directory = get_application_directory()
        default_source = application_directory / "logo_source.pdf"
        default_target = application_directory / "target_form.pdf"
        default_output = application_directory / "output_merged.pdf"

        self.source_path_var = tk.StringVar(
            value=str(default_source) if default_source.exists() else ""
        )
        self.target_path_var = tk.StringVar(
            value=str(default_target) if default_target.exists() else ""
        )
        self.output_path_var = tk.StringVar(value=str(default_output))
        self.source_page_var = tk.StringVar(value="1")
        self.target_pages_var = tk.StringVar(value="すべて")
        self.status_var = tk.StringVar(value="PDFと座標を指定してください。")
        self.last_output_path: Path | None = None
        self.source_coord_vars = [
            tk.StringVar(value=f"{value:g}") for value in DEFAULT_SOURCE_COORDS
        ]
        self.destination_coord_vars = [
            tk.StringVar(value=f"{value:g}")
            for value in DEFAULT_DESTINATION_COORDS
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
            text="抽出元PDFの指定領域を、合成先PDFへベクターのまま配置します。",
        ).grid(row=1, column=0, sticky="w", pady=(0, 12))

        files_frame = ttk.LabelFrame(outer, text="1. ファイル", padding=10)
        files_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        files_frame.columnconfigure(1, weight=1)
        self._add_file_row(files_frame, 0, "抽出元PDF", self.source_path_var, self._select_source_pdf)
        self._add_file_row(files_frame, 1, "合成先PDF", self.target_path_var, self._select_target_pdf)
        self._add_file_row(files_frame, 2, "出力PDF", self.output_path_var, self._select_output_pdf)

        source_frame = ttk.LabelFrame(outer, text="2. 切り取り設定（抽出元）", padding=10)
        source_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(source_frame, text="抽出元ページ").grid(
            row=0, column=0, sticky="w", padx=(0, 6), pady=4
        )
        ttk.Entry(source_frame, textvariable=self.source_page_var, width=8).grid(
            row=0, column=1, sticky="w", pady=4
        )
        ttk.Label(source_frame, text="（1始まり）").grid(
            row=0, column=2, sticky="w", padx=(4, 14), pady=4
        )
        self._add_coordinate_row(source_frame, 1, "切り取り座標", self.source_coord_vars)
        ttk.Button(
            source_frame,
            text="切り取りプレビュー",
            command=self._show_clip_preview,
        ).grid(row=2, column=0, columnspan=6, sticky="w", pady=(8, 2))

        destination_frame = ttk.LabelFrame(
            outer, text="3. 貼り付け設定（合成先）", padding=10
        )
        destination_frame.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        self._add_coordinate_row(
            destination_frame, 0, "貼り付け座標", self.destination_coord_vars
        )
        ttk.Label(destination_frame, text="対象ページ").grid(
            row=1, column=0, sticky="w", padx=(0, 6), pady=4
        )
        ttk.Entry(
            destination_frame, textvariable=self.target_pages_var, width=20
        ).grid(row=1, column=1, columnspan=2, sticky="w", pady=4)
        ttk.Label(destination_frame, text="例: すべて / 1 / 1,3-5").grid(
            row=1, column=3, columnspan=3, sticky="w", padx=(8, 0), pady=4
        )

        ttk.Label(
            outer,
            text=(
                "座標単位は point、原点はページ左上です。"
                "1 mm ≒ 2.83465 point。プレビューで切り取り内容を確認してください。"
            ),
            foreground="#555555",
            wraplength=760,
        ).grid(row=5, column=0, sticky="w", pady=(0, 10))

        action_frame = ttk.Frame(outer)
        action_frame.grid(row=6, column=0, sticky="ew", pady=(2, 10))
        self.run_button = ttk.Button(
            action_frame, text="PDFを合成する", command=self._start_merge
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

    def _select_source_pdf(self) -> None:
        selected = filedialog.askopenfilename(
            title="抽出元PDFを選択",
            initialdir=self._dialog_initial_directory(self.source_path_var.get()),
            filetypes=PDF_FILE_TYPES,
        )
        if selected:
            self.source_path_var.set(selected)
            self.status_var.set("抽出元PDFを選択しました。プレビューを確認してください。")

    def _select_target_pdf(self) -> None:
        selected = filedialog.askopenfilename(
            title="合成先PDFを選択",
            initialdir=self._dialog_initial_directory(self.target_path_var.get()),
            filetypes=PDF_FILE_TYPES,
        )
        if selected:
            self.target_path_var.set(selected)
            current_output = self.output_path_var.get().strip()
            default_name = get_application_directory() / "output_merged.pdf"
            if not current_output or Path(current_output) == default_name:
                target_path = Path(selected)
                self.output_path_var.set(
                    str(target_path.with_name(f"{target_path.stem}_merged.pdf"))
                )
            self.status_var.set("合成先PDFを選択しました。")

    def _select_output_pdf(self) -> None:
        selected = filedialog.asksaveasfilename(
            title="出力PDFを指定",
            initialdir=self._dialog_initial_directory(self.output_path_var.get()),
            initialfile=(
                Path(self.output_path_var.get()).name
                if self.output_path_var.get().strip()
                else "output_merged.pdf"
            ),
            defaultextension=".pdf",
            filetypes=PDF_FILE_TYPES,
        )
        if selected:
            self.output_path_var.set(selected)
            self.status_var.set("出力先を指定しました。")

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
        int,
        tuple[float, float, float, float],
        tuple[float, float, float, float],
        str,
    ]:
        source_text = self.source_path_var.get().strip()
        target_text = self.target_path_var.get().strip()
        output_text = self.output_path_var.get().strip()
        if not source_text:
            raise PdfMergeError("抽出元PDFを選択してください。")
        if not target_text:
            raise PdfMergeError("合成先PDFを選択してください。")
        if not output_text:
            raise PdfMergeError("出力PDFを指定してください。")
        return (
            Path(source_text),
            Path(target_text),
            Path(output_text),
            self._parse_positive_integer(self.source_page_var.get(), "抽出元ページ"),
            self._parse_coordinates(self.source_coord_vars, "切り取り座標"),
            self._parse_coordinates(self.destination_coord_vars, "貼り付け座標"),
            self.target_pages_var.get().strip(),
        )

    def _show_clip_preview(self) -> None:
        try:
            source_text = self.source_path_var.get().strip()
            if not source_text:
                raise PdfMergeError("抽出元PDFを選択してください。")
            page_number = self._parse_positive_integer(
                self.source_page_var.get(), "抽出元ページ"
            )
            coordinates = self._parse_coordinates(
                self.source_coord_vars, "切り取り座標"
            )
            png_bytes, description = create_clip_preview_png(
                Path(source_text), page_number, coordinates
            )
        except Exception as exc:
            self._show_error(exc)
            return

        preview_window = tk.Toplevel(self.root)
        preview_window.title("切り取りプレビュー")
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
            text="このプレビューに見えている内容がPDFへ合成されます。",
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
        self.status_var.set("PDF合成処理を開始します...")
        threading.Thread(
            target=self._merge_worker, args=(inputs,), daemon=True
        ).start()

    def _merge_worker(
        self,
        inputs: tuple[
            Path,
            Path,
            Path,
            int,
            tuple[float, float, float, float],
            tuple[float, float, float, float],
            str,
        ],
    ) -> None:
        try:
            page_count = merge_pdf_area(
                source_path=inputs[0],
                target_path=inputs[1],
                output_path=inputs[2],
                source_page_number=inputs[3],
                source_clip_coords=inputs[4],
                destination_coords=inputs[5],
                target_page_specification=inputs[6],
                progress_callback=self._post_status,
            )
        except Exception as exc:
            self.root.after(0, self._merge_failed, exc)
            return
        self.root.after(0, self._merge_succeeded, inputs[2], page_count)

    def _post_status(self, message: str) -> None:
        self.root.after(0, self.status_var.set, message)

    def _merge_succeeded(self, output_path: Path, page_count: int) -> None:
        self.last_output_path = output_path
        self.run_button.configure(state="normal")
        self.open_button.configure(state="normal")
        self.status_var.set(f"完了: {page_count}ページへ合成しました。\n{output_path}")
        messagebox.showinfo(
            APP_TITLE,
            f"PDF合成が完了しました。\n\n対象ページ数: {page_count}\n{output_path}",
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