"""
Google Slides + Sheets Generator — Creates Google Slides presentations
with charts dynamically linked to Google Sheets data.

Uses the same JSON input format as IBPresentationGenerator but outputs
Google Slides + Sheets URLs instead of a .pptx file.
"""

import os
import re
import logging
from typing import Any

from google.oauth2 import service_account
from googleapiclient.discovery import build

logger = logging.getLogger(__name__)

_EMU_PER_INCH = 914400
_EMU_PER_PT = 12700


def _inches(n: float) -> int:
    return int(n * _EMU_PER_INCH)


def _pt(n: float) -> int:
    return int(n * _EMU_PER_PT)


def _rgb(r: int, g: int, b: int) -> dict:
    """0-255 RGB → Google API color dict (0.0–1.0 floats)."""
    return {"red": r / 255.0, "green": g / 255.0, "blue": b / 255.0}


class GoogleSlidesGenerator:
    """Generates Google Slides presentations with charts linked to Sheets."""

    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/presentations",
        "https://www.googleapis.com/auth/drive",
    ]

    # IB palette
    IB_BLUE = _rgb(0x0F, 0x4D, 0x88)
    DARK_TEXT = _rgb(0x26, 0x26, 0x26)
    TITLE_GRAY = _rgb(0x7F, 0x7F, 0x7F)
    WHITE = _rgb(0xFF, 0xFF, 0xFF)
    BG_LIGHT = _rgb(0xF5, 0xF5, 0xF5)
    SHAPE_DARK = _rgb(0x3C, 0x3C, 0x3C)
    BORDER_GRAY = _rgb(0xD9, 0xDA, 0xDD)
    ACCENT_GRAY = _rgb(0xB3, 0xB5, 0xBB)
    RED_ACCENT = _rgb(0x98, 0x00, 0x00)
    GREEN_ACCENT = _rgb(0x1E, 0x82, 0x32)
    CHART_COLORS = [
        _rgb(0x0F, 0x4D, 0x88),
        _rgb(0x3C, 0x3C, 0x3C),
        _rgb(0xB3, 0xB5, 0xBB),
        _rgb(0x58, 0x59, 0x5C),
    ]

    FONT_TITLE = "Inter"
    FONT_BODY = "Inter"

    SLIDE_W = _inches(13.333)
    SLIDE_H = _inches(7.5)
    MARGIN_L = _inches(0.6)
    CONTENT_W = _inches(12.133)
    TITLE_TOP = _inches(0.35)
    TITLE_H = _inches(0.55)
    SEPARATOR_TOP = _inches(0.95)
    CONTENT_TOP = _inches(1.2)
    FOOTER_TOP = _inches(7.1)

    def __init__(self, credentials_path: str | None = None, folder_id: str | None = None):
        creds_path = credentials_path or os.environ.get(
            "GOOGLE_APPLICATION_CREDENTIALS"
        )
        if not creds_path or not os.path.exists(creds_path):
            raise ValueError(
                "Google credentials file not found. "
                "Set GOOGLE_APPLICATION_CREDENTIALS or pass credentials_path."
            )

        self._credentials = service_account.Credentials.from_service_account_file(
            creds_path, scopes=self.SCOPES
        )
        self._sheets = build("sheets", "v4", credentials=self._credentials)
        self._slides = build("slides", "v1", credentials=self._credentials)
        self._drive = build("drive", "v3", credentials=self._credentials)
        self._folder_id = folder_id or os.environ.get("DRIVE_FOLDER_ID")
        self._page_num = 0

    # ================================================================ #
    #  PUBLIC                                                           #
    # ================================================================ #

    def generate(self, data: dict) -> dict:
        """Generate a Google Slides + Sheets pair and return their URLs."""
        slides = data.get("slides", [])
        if not slides:
            raise ValueError("No slides provided")

        title = data.get("presentation_title", "Apresentação")

        chart_slides = [
            (i, s) for i, s in enumerate(slides) if s.get("layout") == "chart"
        ]

        spreadsheet_id = None
        chart_map: dict[int, int] = {}

        if chart_slides:
            spreadsheet_id, chart_map = self._create_spreadsheet(
                title, chart_slides
            )

        presentation_id = self._create_presentation(
            title, slides, spreadsheet_id, chart_map
        )

        self._set_permissions(presentation_id)
        if spreadsheet_id:
            self._set_permissions(spreadsheet_id)

        slides_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
        sheets_url = (
            f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
            if spreadsheet_id
            else None
        )

        return {
            "slides_url": slides_url,
            "sheets_url": sheets_url,
            "presentation_title": title,
        }

    # ================================================================ #
    #  SPREADSHEET CREATION                                             #
    # ================================================================ #

    def _create_spreadsheet(
        self, title: str, chart_slides: list[tuple[int, dict]]
    ) -> tuple[str, dict]:
        sheet_names: list[str] = []
        for idx, (_, slide_data) in enumerate(chart_slides):
            raw_name = slide_data.get("title", f"Chart {idx + 1}")
            safe_name = re.sub(r"[^\w\s-]", "", raw_name)[:30].strip() or f"Chart_{idx + 1}"
            sheet_names.append(safe_name)

        file_meta: dict[str, Any] = {
            "name": f"{title} - Dados",
            "mimeType": "application/vnd.google-apps.spreadsheet",
        }
        if self._folder_id:
            file_meta["parents"] = [self._folder_id]

        created = self._drive.files().create(body=file_meta, fields="id").execute()
        spreadsheet_id: str = created["id"]

        ss = self._sheets.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        default_sheet_id = ss["sheets"][0]["properties"]["sheetId"]

        batch_reqs: list[dict] = []
        sheet_id_map: list[int] = []

        for idx, name in enumerate(sheet_names):
            if idx == 0:
                batch_reqs.append({
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": default_sheet_id,
                            "title": name,
                            "index": 0,
                        },
                        "fields": "title,index",
                    }
                })
                sheet_id_map.append(default_sheet_id)
            else:
                new_id = idx * 100
                batch_reqs.append({
                    "addSheet": {
                        "properties": {
                            "sheetId": new_id,
                            "title": name,
                            "index": idx,
                        }
                    }
                })
                sheet_id_map.append(new_id)

        if batch_reqs:
            self._sheets.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": batch_reqs},
            ).execute()

        chart_map: dict[int, int] = {}

        for idx, (slide_idx, slide_data) in enumerate(chart_slides):
            cd = slide_data.get("chart_data") or {}
            categories = cd.get("categories", [])
            series_list = cd.get("series", [])
            chart_type_str = slide_data.get("chart_type", "bar")
            chart_title = slide_data.get("title", f"Chart {idx + 1}")

            if not categories or not series_list:
                continue

            sheet_id = sheet_id_map[idx]

            self._populate_sheet(
                spreadsheet_id, sheet_names[idx], categories, series_list
            )
            chart_id = self._add_chart_to_sheet(
                spreadsheet_id,
                sheet_id,
                chart_title,
                chart_type_str,
                len(categories),
                series_list,
            )
            if chart_id is not None:
                chart_map[slide_idx] = chart_id

        return spreadsheet_id, chart_map

    def _populate_sheet(
        self,
        spreadsheet_id: str,
        sheet_name: str,
        categories: list[str],
        series_list: list[dict],
    ) -> None:
        header = [""] + [
            s.get("name", f"Series {j + 1}") for j, s in enumerate(series_list)
        ]
        rows: list[list[Any]] = [header]
        for ci, cat in enumerate(categories):
            row: list[Any] = [str(cat)]
            for s in series_list:
                vals = s.get("values", [])
                row.append(vals[ci] if ci < len(vals) else 0)
            rows.append(row)

        last_col = chr(65 + len(series_list))
        range_a1 = f"'{sheet_name}'!A1:{last_col}{len(rows)}"

        self._sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range_a1,
            valueInputOption="RAW",
            body={"values": rows},
        ).execute()

    def _add_chart_to_sheet(
        self,
        spreadsheet_id: str,
        sheet_id: int,
        chart_title: str,
        chart_type_str: str,
        n_categories: int,
        series_list: list[dict],
    ) -> int | None:
        type_map = {
            "bar": "COLUMN",
            "line": "LINE",
            "pie": "PIE",
            "stacked_bar": "COLUMN",
        }
        sheets_type = type_map.get(chart_type_str, "COLUMN")
        is_stacked = chart_type_str == "stacked_bar"
        n_rows = n_categories + 1
        n_series = len(series_list)

        chart_spec: dict[str, Any] = {
            "title": chart_title,
            "titleTextFormat": {
                "fontFamily": self.FONT_TITLE,
                "fontSize": 14,
                "bold": True,
                "foregroundColorStyle": {"rgbColor": self.DARK_TEXT},
            },
        }

        if sheets_type == "PIE":
            chart_spec["pieChart"] = {
                "legendPosition": "LABELED_LEGEND",
                "domain": {
                    "sourceRange": {
                        "sources": [
                            {
                                "sheetId": sheet_id,
                                "startRowIndex": 0,
                                "endRowIndex": n_rows,
                                "startColumnIndex": 0,
                                "endColumnIndex": 1,
                            }
                        ]
                    }
                },
                "series": {
                    "sourceRange": {
                        "sources": [
                            {
                                "sheetId": sheet_id,
                                "startRowIndex": 0,
                                "endRowIndex": n_rows,
                                "startColumnIndex": 1,
                                "endColumnIndex": 2,
                            }
                        ]
                    }
                },
            }
        else:
            series_spec = []
            for si in range(n_series):
                col = si + 1
                entry: dict[str, Any] = {
                    "series": {
                        "sourceRange": {
                            "sources": [
                                {
                                    "sheetId": sheet_id,
                                    "startRowIndex": 0,
                                    "endRowIndex": n_rows,
                                    "startColumnIndex": col,
                                    "endColumnIndex": col + 1,
                                }
                            ]
                        }
                    },
                    "targetAxis": "LEFT_AXIS",
                }
                if si < len(self.CHART_COLORS):
                    entry["colorStyle"] = {"rgbColor": self.CHART_COLORS[si]}
                series_spec.append(entry)

            basic: dict[str, Any] = {
                "chartType": sheets_type,
                "legendPosition": (
                    "BOTTOM_LEGEND" if n_series > 1 else "NO_LEGEND"
                ),
                "headerCount": 1,
                "domains": [
                    {
                        "domain": {
                            "sourceRange": {
                                "sources": [
                                    {
                                        "sheetId": sheet_id,
                                        "startRowIndex": 0,
                                        "endRowIndex": n_rows,
                                        "startColumnIndex": 0,
                                        "endColumnIndex": 1,
                                    }
                                ]
                            }
                        }
                    }
                ],
                "series": series_spec,
                "axis": [
                    {
                        "position": "BOTTOM_AXIS",
                        "format": {
                            "fontFamily": self.FONT_BODY,
                            "fontSize": 10,
                        },
                    },
                    {
                        "position": "LEFT_AXIS",
                        "format": {
                            "fontFamily": self.FONT_BODY,
                            "fontSize": 10,
                        },
                    },
                ],
            }
            if is_stacked:
                basic["stackedType"] = "STACKED"
            chart_spec["basicChart"] = basic

        resp = (
            self._sheets.spreadsheets()
            .batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={
                    "requests": [
                        {
                            "addChart": {
                                "chart": {
                                    "spec": chart_spec,
                                    "position": {
                                        "overlayPosition": {
                                            "anchorCell": {
                                                "sheetId": sheet_id,
                                                "rowIndex": n_rows + 2,
                                                "columnIndex": 0,
                                            },
                                            "widthPixels": 800,
                                            "heightPixels": 450,
                                        }
                                    },
                                }
                            }
                        }
                    ]
                },
            )
            .execute()
        )

        for reply in resp.get("replies", []):
            if "addChart" in reply:
                return reply["addChart"]["chart"]["chartId"]
        return None

    # ================================================================ #
    #  PRESENTATION CREATION                                            #
    # ================================================================ #

    def _create_presentation(
        self,
        title: str,
        slides: list[dict],
        spreadsheet_id: str | None,
        chart_map: dict[int, int],
    ) -> str:
        file_meta: dict[str, Any] = {
            "name": title,
            "mimeType": "application/vnd.google-apps.presentation",
        }
        if self._folder_id:
            file_meta["parents"] = [self._folder_id]

        created = self._drive.files().create(body=file_meta, fields="id").execute()
        presentation_id: str = created["id"]

        pres = self._slides.presentations().get(
            presentationId=presentation_id
        ).execute()
        default_slide_id: str = pres["slides"][0]["objectId"]

        actual_w = pres.get("pageSize", {}).get("width", {}).get("magnitude", self.SLIDE_W)
        actual_h = pres.get("pageSize", {}).get("height", {}).get("magnitude", self.SLIDE_H)
        design_w = type(self).SLIDE_W
        design_h = type(self).SLIDE_H
        if actual_w != design_w or actual_h != design_h:
            sx = actual_w / design_w
            sy = actual_h / design_h
            self.SLIDE_W = int(actual_w)
            self.SLIDE_H = int(actual_h)
            self.MARGIN_L = int(type(self).MARGIN_L * sx)
            self.CONTENT_W = int(type(self).CONTENT_W * sx)
            self.TITLE_TOP = int(type(self).TITLE_TOP * sy)
            self.TITLE_H = int(type(self).TITLE_H * sy)
            self.SEPARATOR_TOP = int(type(self).SEPARATOR_TOP * sy)
            self.CONTENT_TOP = int(type(self).CONTENT_TOP * sy)
            self.FOOTER_TOP = int(type(self).FOOTER_TOP * sy)

        setup_requests: list[dict] = [
            {"deleteObject": {"objectId": default_slide_id}}
        ]
        slide_ids: list[str] = []
        for i in range(len(slides)):
            sid = f"slide_{i}"
            slide_ids.append(sid)
            setup_requests.append(
                {"createSlide": {"objectId": sid, "insertionIndex": i}}
            )

        self._slides.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": setup_requests},
        ).execute()

        self._page_num = 0
        for i, slide_data in enumerate(slides):
            self._page_num += 1
            layout = slide_data.get("layout", "content")
            builder = {
                "cover": self._build_cover,
                "content": self._build_content,
                "two_columns": self._build_two_columns,
                "table": self._build_table,
                "chart": self._build_chart,
                "key_metrics": self._build_key_metrics,
                "closing": self._build_closing,
            }.get(layout, self._build_content)

            if layout == "chart":
                reqs = builder(slide_ids[i], slide_data, i, spreadsheet_id, chart_map)
            else:
                reqs = builder(slide_ids[i], slide_data)

            if reqs:
                self._slides.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={"requests": reqs},
                ).execute()

        return presentation_id

    # ================================================================ #
    #  SLIDE BUILDERS                                                   #
    # ================================================================ #

    def _build_cover(self, sid: str, data: dict) -> list[dict]:
        reqs = [self._bg(sid, self.SHAPE_DARK)]

        title = data.get("title", "")
        if title:
            reqs += self._textbox(
                sid, f"{sid}_t", title,
                self.MARGIN_L, _inches(2.4), self.CONTENT_W, _inches(1.6),
                size=40, color=self.WHITE, bold=True, font=self.FONT_TITLE,
            )

        subtitle = data.get("subtitle", "")
        if subtitle:
            reqs += self._textbox(
                sid, f"{sid}_s", subtitle,
                self.MARGIN_L, _inches(4.2), self.CONTENT_W, _inches(0.9),
                size=18, color=self.ACCENT_GRAY,
            )

        date = data.get("date", "")
        if date:
            reqs += self._textbox(
                sid, f"{sid}_d", date,
                self.MARGIN_L, _inches(6.5), self.CONTENT_W, _inches(0.45),
                size=14, color=self.TITLE_GRAY,
            )
        return reqs

    def _build_content(self, sid: str, data: dict) -> list[dict]:
        reqs = [self._bg(sid, self.WHITE)]
        reqs += self._title_area(sid, data.get("title", ""))
        reqs += self._footer(sid)

        bullets = data.get("bullets", [])
        if bullets:
            text = "\n".join(f"\u2022  {b}" for b in bullets)
            reqs += self._textbox(
                sid, f"{sid}_bul", text,
                self.MARGIN_L + _inches(0.3),
                self.CONTENT_TOP + _inches(0.15),
                self.CONTENT_W - _inches(0.6),
                _inches(5.0),
                size=14, color=self.DARK_TEXT, line_spacing=200,
            )
        return reqs

    def _build_two_columns(self, sid: str, data: dict) -> list[dict]:
        reqs = [self._bg(sid, self.WHITE)]
        reqs += self._title_area(sid, data.get("title", ""))
        reqs += self._footer(sid)

        col_w = _inches(5.6)
        gap = _inches(0.9)

        for ci, key in enumerate(("left_column", "right_column")):
            col = data.get(key) or {}
            left = self.MARGIN_L + ci * (col_w + gap)

            sub = col.get("subtitle", "")
            if sub:
                reqs += self._textbox(
                    sid, f"{sid}_c{ci}s", sub,
                    left, self.CONTENT_TOP, col_w, _inches(0.42),
                    size=16, color=self.IB_BLUE, bold=True, font=self.FONT_TITLE,
                )
                reqs += self._rect(
                    sid, f"{sid}_c{ci}l",
                    left, self.CONTENT_TOP + _inches(0.48), col_w, _inches(0.015),
                    fill=self.BORDER_GRAY,
                )

            bul = col.get("bullets", [])
            if bul:
                text = "\n".join(f"\u2022  {b}" for b in bul)
                reqs += self._textbox(
                    sid, f"{sid}_c{ci}b", text,
                    left + _inches(0.15),
                    self.CONTENT_TOP + _inches(0.65),
                    col_w - _inches(0.3),
                    _inches(4.5),
                    size=13, color=self.DARK_TEXT, line_spacing=185,
                )

        reqs += self._rect(
            sid, f"{sid}_div",
            self.MARGIN_L + col_w + gap // 2,
            self.CONTENT_TOP,
            _inches(0.016), _inches(5.2),
            fill=self.BORDER_GRAY,
        )
        return reqs

    def _build_table(self, sid: str, data: dict) -> list[dict]:
        reqs = [self._bg(sid, self.WHITE)]
        reqs += self._title_area(sid, data.get("title", ""))
        reqs += self._footer(sid)

        tbl = data.get("table") or {}
        headers = tbl.get("headers", [])
        rows = tbl.get("rows", [])
        if not headers:
            return reqs

        n_rows = len(rows) + 1
        n_cols = len(headers)
        table_id = f"{sid}_tbl"

        reqs.append({
            "createTable": {
                "objectId": table_id,
                "elementProperties": {
                    "pageObjectId": sid,
                    "size": {
                        "width": {"magnitude": self.CONTENT_W, "unit": "EMU"},
                        "height": {
                            "magnitude": _inches(min(0.42 * n_rows, 5.2)),
                            "unit": "EMU",
                        },
                    },
                    "transform": {
                        "scaleX": 1, "scaleY": 1,
                        "translateX": self.MARGIN_L,
                        "translateY": self.CONTENT_TOP + _inches(0.15),
                        "unit": "EMU",
                    },
                },
                "rows": n_rows,
                "columns": n_cols,
            }
        })

        for ci, h in enumerate(headers):
            reqs.append({
                "insertText": {
                    "objectId": table_id,
                    "cellLocation": {"rowIndex": 0, "columnIndex": ci},
                    "text": str(h),
                    "insertionIndex": 0,
                }
            })

        for ri, row in enumerate(rows):
            for ci, val in enumerate(row):
                if ci >= n_cols:
                    break
                reqs.append({
                    "insertText": {
                        "objectId": table_id,
                        "cellLocation": {"rowIndex": ri + 1, "columnIndex": ci},
                        "text": str(val),
                        "insertionIndex": 0,
                    }
                })

        for ci in range(n_cols):
            reqs.append({
                "updateTableCellProperties": {
                    "objectId": table_id,
                    "tableRange": {
                        "location": {"rowIndex": 0, "columnIndex": ci},
                        "rowSpan": 1, "columnSpan": 1,
                    },
                    "tableCellProperties": {
                        "tableCellBackgroundFill": {
                            "solidFill": {"color": {"rgbColor": self.IB_BLUE}}
                        }
                    },
                    "fields": "tableCellBackgroundFill",
                }
            })
            reqs.append({
                "updateTextStyle": {
                    "objectId": table_id,
                    "cellLocation": {"rowIndex": 0, "columnIndex": ci},
                    "style": {
                        "foregroundColor": {"opaqueColor": {"rgbColor": self.WHITE}},
                        "fontFamily": self.FONT_BODY,
                        "fontSize": {"magnitude": 12, "unit": "PT"},
                        "bold": True,
                    },
                    "textRange": {"type": "ALL"},
                    "fields": "foregroundColor,fontFamily,fontSize,bold",
                }
            })

        for ri, row in enumerate(rows):
            bg = self.BG_LIGHT if ri % 2 == 0 else self.WHITE
            for ci in range(min(len(row), n_cols)):
                reqs.append({
                    "updateTableCellProperties": {
                        "objectId": table_id,
                        "tableRange": {
                            "location": {"rowIndex": ri + 1, "columnIndex": ci},
                            "rowSpan": 1, "columnSpan": 1,
                        },
                        "tableCellProperties": {
                            "tableCellBackgroundFill": {
                                "solidFill": {"color": {"rgbColor": bg}}
                            }
                        },
                        "fields": "tableCellBackgroundFill",
                    }
                })
                val = str(row[ci])
                tc = self.DARK_TEXT
                if val.startswith("+"):
                    tc = self.GREEN_ACCENT
                elif val.startswith("-") and "%" in val:
                    tc = self.RED_ACCENT
                reqs.append({
                    "updateTextStyle": {
                        "objectId": table_id,
                        "cellLocation": {"rowIndex": ri + 1, "columnIndex": ci},
                        "style": {
                            "foregroundColor": {"opaqueColor": {"rgbColor": tc}},
                            "fontFamily": self.FONT_BODY,
                            "fontSize": {"magnitude": 11, "unit": "PT"},
                        },
                        "textRange": {"type": "ALL"},
                        "fields": "foregroundColor,fontFamily,fontSize",
                    }
                })
        return reqs

    def _build_chart(
        self,
        sid: str,
        data: dict,
        slide_index: int,
        spreadsheet_id: str | None,
        chart_map: dict[int, int],
    ) -> list[dict]:
        reqs = [self._bg(sid, self.WHITE)]
        reqs += self._title_area(sid, data.get("title", ""))
        reqs += self._footer(sid)

        chart_id = chart_map.get(slide_index)
        if not spreadsheet_id or chart_id is None:
            return reqs

        reqs.append({
            "createSheetsChart": {
                "objectId": f"{sid}_chart",
                "spreadsheetId": spreadsheet_id,
                "chartId": chart_id,
                "linkingMode": "LINKED",
                "elementProperties": {
                    "pageObjectId": sid,
                    "size": {
                        "width": {
                            "magnitude": self.CONTENT_W - _inches(1.0),
                            "unit": "EMU",
                        },
                        "height": {"magnitude": _inches(4.8), "unit": "EMU"},
                    },
                    "transform": {
                        "scaleX": 1, "scaleY": 1,
                        "translateX": self.MARGIN_L + _inches(0.5),
                        "translateY": self.CONTENT_TOP + _inches(0.2),
                        "unit": "EMU",
                    },
                },
            }
        })
        return reqs

    def _build_key_metrics(self, sid: str, data: dict) -> list[dict]:
        reqs = [self._bg(sid, self.WHITE)]
        reqs += self._title_area(sid, data.get("title", ""))
        reqs += self._footer(sid)

        metrics = data.get("metrics", [])
        if not metrics:
            return reqs

        n = len(metrics)
        cols = min(n, 4)
        gap = _inches(0.35)
        card_w = (self.CONTENT_W - gap * (cols - 1)) // cols
        card_h = _inches(2.2)
        start_y = self.CONTENT_TOP + _inches(0.6)

        for i, m in enumerate(metrics):
            r = i // cols
            c = i % cols
            left = self.MARGIN_L + c * (card_w + gap)
            top = start_y + r * (card_h + gap)

            reqs += self._rect(
                sid, f"{sid}_cd{i}",
                left, top, card_w, card_h,
                fill=self.BG_LIGHT, outline=self.BORDER_GRAY,
                shape="ROUND_RECTANGLE",
            )
            reqs += self._rect(
                sid, f"{sid}_cb{i}",
                left, top, card_w, _inches(0.04),
                fill=self.IB_BLUE,
            )

            label = m.get("label", "")
            if label:
                reqs += self._textbox(
                    sid, f"{sid}_ml{i}", label,
                    left + _inches(0.2), top + _inches(0.25),
                    card_w - _inches(0.4), _inches(0.35),
                    size=12, color=self.TITLE_GRAY, align="CENTER",
                )

            value = m.get("value", "")
            if value:
                reqs += self._textbox(
                    sid, f"{sid}_mv{i}", value,
                    left + _inches(0.2), top + _inches(0.7),
                    card_w - _inches(0.4), _inches(0.65),
                    size=26, color=self.DARK_TEXT, bold=True,
                    font=self.FONT_TITLE, align="CENTER",
                )

            variation = m.get("variation", "")
            if variation:
                vc = self.TITLE_GRAY
                if "+" in variation:
                    vc = self.GREEN_ACCENT
                elif "-" in variation:
                    vc = self.RED_ACCENT
                reqs += self._textbox(
                    sid, f"{sid}_mx{i}", variation,
                    left + _inches(0.2), top + _inches(1.5),
                    card_w - _inches(0.4), _inches(0.35),
                    size=13, color=vc, bold=True, align="CENTER",
                )
        return reqs

    def _build_closing(self, sid: str, data: dict) -> list[dict]:
        reqs = [self._bg(sid, self.SHAPE_DARK)]

        title = data.get("title", "Obrigado")
        reqs += self._textbox(
            sid, f"{sid}_t", title,
            self.MARGIN_L, _inches(2.8), self.CONTENT_W, _inches(1.2),
            size=42, color=self.WHITE, bold=True,
            font=self.FONT_TITLE, align="CENTER",
        )

        subtitle = data.get("subtitle", "")
        if subtitle:
            reqs += self._textbox(
                sid, f"{sid}_s", subtitle,
                self.MARGIN_L, _inches(4.2), self.CONTENT_W, _inches(0.7),
                size=16, color=self.ACCENT_GRAY, align="CENTER",
            )
        return reqs

    # ================================================================ #
    #  REQUEST HELPERS                                                  #
    # ================================================================ #

    @staticmethod
    def _bg(slide_id: str, color: dict) -> dict:
        return {
            "updatePageProperties": {
                "objectId": slide_id,
                "pageProperties": {
                    "pageBackgroundFill": {
                        "solidFill": {"color": {"rgbColor": color}}
                    }
                },
                "fields": "pageBackgroundFill",
            }
        }

    def _textbox(
        self,
        slide_id: str,
        eid: str,
        text: str,
        left: int,
        top: int,
        width: int,
        height: int,
        *,
        font: str | None = None,
        size: int = 14,
        color: dict | None = None,
        bold: bool = False,
        align: str = "START",
        line_spacing: int | None = None,
    ) -> list[dict]:
        reqs: list[dict] = [
            {
                "createShape": {
                    "objectId": eid,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {
                            "width": {"magnitude": width, "unit": "EMU"},
                            "height": {"magnitude": height, "unit": "EMU"},
                        },
                        "transform": {
                            "scaleX": 1, "scaleY": 1,
                            "translateX": left, "translateY": top,
                            "unit": "EMU",
                        },
                    },
                }
            },
            {
                "insertText": {
                    "objectId": eid,
                    "text": text,
                    "insertionIndex": 0,
                }
            },
        ]

        style: dict[str, Any] = {
            "fontFamily": font or self.FONT_BODY,
            "fontSize": {"magnitude": size, "unit": "PT"},
            "bold": bold,
        }
        fields = "fontFamily,fontSize,bold"
        if color:
            style["foregroundColor"] = {"opaqueColor": {"rgbColor": color}}
            fields += ",foregroundColor"

        reqs.append({
            "updateTextStyle": {
                "objectId": eid,
                "style": style,
                "textRange": {"type": "ALL"},
                "fields": fields,
            }
        })

        pstyle: dict[str, Any] = {"alignment": align}
        pfields = "alignment"
        if line_spacing:
            pstyle["lineSpacing"] = line_spacing
            pfields += ",lineSpacing"
        reqs.append({
            "updateParagraphStyle": {
                "objectId": eid,
                "style": pstyle,
                "textRange": {"type": "ALL"},
                "fields": pfields,
            }
        })
        return reqs

    def _rect(
        self,
        slide_id: str,
        eid: str,
        left: int,
        top: int,
        width: int,
        height: int,
        *,
        fill: dict | None = None,
        outline: dict | None = None,
        shape: str = "RECTANGLE",
    ) -> list[dict]:
        reqs: list[dict] = [
            {
                "createShape": {
                    "objectId": eid,
                    "shapeType": shape,
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {
                            "width": {"magnitude": width, "unit": "EMU"},
                            "height": {"magnitude": height, "unit": "EMU"},
                        },
                        "transform": {
                            "scaleX": 1, "scaleY": 1,
                            "translateX": left, "translateY": top,
                            "unit": "EMU",
                        },
                    },
                }
            },
        ]

        props: dict[str, Any] = {}
        fields: list[str] = []

        if fill:
            props["shapeBackgroundFill"] = {
                "solidFill": {"color": {"rgbColor": fill}, "alpha": 1.0}
            }
            fields.append("shapeBackgroundFill")

        if outline:
            props["outline"] = {
                "outlineFill": {
                    "solidFill": {"color": {"rgbColor": outline}}
                },
                "weight": {"magnitude": 1, "unit": "PT"},
            }
            fields.append("outline")
        else:
            props["outline"] = {"propertyState": "NOT_RENDERED"}
            fields.append("outline")

        if fields:
            reqs.append({
                "updateShapeProperties": {
                    "objectId": eid,
                    "shapeProperties": props,
                    "fields": ",".join(fields),
                }
            })
        return reqs

    def _title_area(self, sid: str, title: str) -> list[dict]:
        reqs: list[dict] = []
        if title:
            reqs += self._textbox(
                sid, f"{sid}_tt", title,
                self.MARGIN_L, self.TITLE_TOP, self.CONTENT_W, self.TITLE_H,
                size=24, color=self.TITLE_GRAY, bold=True, font=self.FONT_TITLE,
            )
        reqs += self._rect(
            sid, f"{sid}_sep",
            self.MARGIN_L, self.SEPARATOR_TOP,
            self.CONTENT_W, _inches(0.015),
            fill=self.BORDER_GRAY,
        )
        return reqs

    def _footer(self, sid: str) -> list[dict]:
        reqs = self._rect(
            sid, f"{sid}_fl",
            0, self.FOOTER_TOP, self.SLIDE_W, _inches(0.01),
            fill=self.BORDER_GRAY,
        )
        reqs += self._textbox(
            sid, f"{sid}_pn", str(self._page_num),
            self.SLIDE_W - _inches(1.0), self.FOOTER_TOP + _inches(0.04),
            _inches(0.5), _inches(0.3),
            size=10, color=self.TITLE_GRAY, align="END",
        )
        return reqs

    # ================================================================ #
    #  PERMISSIONS                                                      #
    # ================================================================ #

    def _set_permissions(self, file_id: str) -> None:
        share_email = os.environ.get("GOOGLE_SHARE_EMAIL")
        if share_email:
            self._drive.permissions().create(
                fileId=file_id,
                body={
                    "type": "user",
                    "role": "writer",
                    "emailAddress": share_email,
                },
                sendNotificationEmail=False,
            ).execute()

        self._drive.permissions().create(
            fileId=file_id,
            body={"type": "anyone", "role": "writer"},
        ).execute()
