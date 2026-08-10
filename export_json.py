# -*- coding: utf-8 -*-
"""출석부 데이터를 홈페이지 게시용 JSON으로 내보낸다 (읽기 전용).

각 과목(math/korean/english)의 '최신 월' 출석부를 읽어서
data-export/<과목>.json 과 data-export/all.json 을 만든다.
봇 코드·데이터는 건드리지 않는다. 주기 실행(launchd/cron)으로 최신 유지.

JSON 구조:
{
  "subject": "math", "display": "Zest 수학과", "month": "26.08",
  "updated": "2026-08-07T22:00:00+09:00",
  "classes": {
    "중3": {
      "students": ["김유현", ...],
      "dates": {
        "8/3": {
          "label": "8/3(월)", "비고라벨": "비고",
          "출석": {"김유현":"O", "김효담":"X (여행)"},
          "수업내용": {"김유현":"원주각", ...},   # 반 전체 동일하면 모두 같은 값
          "과제수행": {...}, "다음과제": {...}, "비고": {...}
        }, ...
      }
    }, ...
  }
}
"""
import os
import re
import json
import datetime
from zoneinfo import ZoneInfo

import attendance_core as ac
import subjects

BASE = os.path.dirname(os.path.abspath(__file__))
KST = ZoneInfo("Asia/Seoul")
OUT_DIR = os.path.join(BASE, "data-export")
SUBJECT_DIRS = {"math": "data-math", "korean": "data-korean", "english": "data-english"}
ROWS = ["출석", "수업내용", "과제수행", "다음과제", "비고"]
_FNAME_RE = re.compile(r"^(\d{2})\.(\d{2}) 출석부\.xlsx$")


def latest_file(dd):
    """폴더에서 가장 최근 'YY.MM 출석부.xlsx' 경로와 'YY.MM'. 없으면 (None, None)."""
    if not os.path.isdir(dd):
        return None, None
    best, bestkey = None, None
    for f in os.listdir(dd):
        m = _FNAME_RE.match(f)
        if m:
            key = (int(m.group(1)), int(m.group(2)))
            if bestkey is None or key > bestkey:
                bestkey, best = key, f
    if not best:
        return None, None
    return os.path.join(dd, best), f"{bestkey[0]:02d}.{bestkey[1]:02d}"


def _merge_anchor(ws):
    """병합 셀의 각 좌표 → (앵커행, 앵커열). 병합 안이면 앵커값을 쓰기 위함."""
    out = {}
    for rng in ws.merged_cells.ranges:
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                out[(r, c)] = (rng.min_row, rng.min_col)
    return out


def export_subject(key, dd):
    path, month = latest_file(dd)
    if not path:
        return None
    wb = ac.load_workbook(path)
    data = {
        "subject": key,
        "display": subjects.get(key)["display"],
        "month": month,
        "updated": datetime.datetime.now(KST).isoformat(timespec="seconds"),
        "classes": {},
    }
    for s in wb.sheetnames:
        ws = wb[s]
        roster = ac.get_roster(ws)
        if not roster:
            continue
        start = ac._find_start_row(ws)
        if start is None:
            continue
        anchor = _merge_anchor(ws)

        def eff(r, c):
            ar, acol = anchor.get((r, c), (r, c))
            return ws.cell(ar, acol).value

        cls = {"students": list(roster.keys()), "dates": {}}
        for top in range(start, ws.max_row + 1, ac.BLOCK_SIZE):
            lbl = ws.cell(top, 1).value
            m = re.search(r"(\d{1,2})[/.](\d{1,2})", str(lbl or ""))
            if not m:
                continue
            ds = f"{int(m.group(1))}/{int(m.group(2))}"
            block = {"label": str(lbl).strip()}
            for i, row in enumerate(ROWS):
                rr = top + i
                if row == "비고":
                    block["비고라벨"] = ws.cell(rr, 2).value or "비고"
                vals = {}
                for nm, col in roster.items():
                    v = eff(rr, col)
                    if v not in (None, ""):
                        vals[nm] = v
                block[row] = vals
            cls["dates"][ds] = block
        data["classes"][s] = cls
    return data


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    combined = {"updated": datetime.datetime.now(KST).isoformat(timespec="seconds"),
                "subjects": {}}
    for key, sub in SUBJECT_DIRS.items():
        data = export_subject(key, os.path.join(BASE, sub))
        if data is None:
            continue
        with open(os.path.join(OUT_DIR, f"{key}.json"), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
        combined["subjects"][key] = data
    with open(os.path.join(OUT_DIR, "all.json"), "w", encoding="utf-8") as f:
        json.dump(combined, f, ensure_ascii=False, indent=1)
    print("export 완료:", ", ".join(combined["subjects"]) or "(데이터 없음)")


if __name__ == "__main__":
    main()
