import re
import json
import warnings
from dataclasses import dataclass, field
from datetime import datetime, timezone, timedelta

import gspread
from google.colab import auth
from google.auth import default
from googleapiclient.discovery import build

warnings.filterwarnings("ignore", module="google_auth_httplib2")

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1vbJOCrOp_NGBKnY5uQHgRUzxiHPi_BaNxkDiHr0W16I/edit"
READABLE_DOC_ID = "16MIqR6XZp9y0S_qoChHSRd0N5qMKdwoRG28w4YhsusU"
KST = timezone(timedelta(hours=9))

APPLY_LOG_SHEET_NAME = "반영로그"
APPLY_LOG_HEADERS = ["No", "시각", "제목", "JSON", "상점현황_JSON", "실행취소"]

_docs_service = None
_creds = None


@dataclass
class ApplyResult:
    """상점 반영·로그 기록의 단계별 성공 여부."""

    tool: str = ""
    timestamp: str = ""
    member_count: int = 0
    sheet_ok: bool = False
    sheet_error: str | None = None
    machine_log_ok: bool = False
    machine_log_no: int | None = None
    machine_log_error: str | None = None
    readable_log_ok: bool = False
    readable_log_error: str | None = None
    warnings: list[str] = field(default_factory=list)

    @property
    def is_full_success(self) -> bool:
        return self.sheet_ok and self.machine_log_ok and self.readable_log_ok

    @property
    def is_dangerous(self) -> bool:
        """시트만 반영되고 기계 로그가 없으면 실행취소 불가."""
        return self.sheet_ok and not self.machine_log_ok

    def merge_log(self, log_result: "ApplyResult") -> None:
        self.machine_log_ok = log_result.machine_log_ok
        self.machine_log_no = log_result.machine_log_no
        self.machine_log_error = log_result.machine_log_error
        self.readable_log_ok = log_result.readable_log_ok
        self.readable_log_error = log_result.readable_log_error
        self.warnings.extend(log_result.warnings)


def print_apply_summary(result: ApplyResult) -> None:
    """반영 결과를 단계별로 출력한다."""
    def _status(ok: bool) -> str:
        return "✅ 성공" if ok else "❌ 실패"

    print("\n" + "=" * 42)
    print("반영 결과 요약")
    print("=" * 42)
    if result.tool:
        print(f"도구     : {result.tool}")
    if result.timestamp:
        print(f"시각     : {result.timestamp}")
    if result.member_count:
        print(f"대상     : {result.member_count}명")
    print(f"[상점현황 시트] {_status(result.sheet_ok)}")
    if result.sheet_error:
        print(f"  └ {result.sheet_error}")
    if result.sheet_ok:
        if result.machine_log_ok and result.machine_log_no:
            print(f"[반영로그]    {_status(True)} → No. {result.machine_log_no} (실행취소용)")
        else:
            print(f"[반영로그]    {_status(False)}")
            if result.machine_log_error:
                print(f"  └ {result.machine_log_error}")
        if result.machine_log_ok:
            print(f"[상점 반영기록]   {_status(result.readable_log_ok)}")
            if result.readable_log_error:
                print(f"  └ {result.readable_log_error}")
        else:
            print("[상점 반영기록]   ⏭️ 건너뜀 (반영로그 없음)")
    for w in result.warnings:
        print(f"⚠️ {w}")
    print("-" * 42)
    if result.is_full_success:
        print("전체 상태: ✅ 완료 (시트 + 로그 모두 반영)")
    elif result.is_dangerous:
        print("전체 상태: ⚠️ 부분 성공 (위험)")
        print("  점수는 시트에 반영되었으나 반영로그 기록이 없습니다.")
        print("  → 실행취소 도구로 복구할 수 없습니다. 수동 조정이 필요합니다.")
    elif not result.sheet_ok:
        print("전체 상태: ❌ 시트 반영 실패")
    else:
        print("전체 상태: ⚠️ 부분 성공")
        print("  → 상점 반영기록만 실패한 경우, 반영로그 No.로 실행취소는 가능합니다.")
    print("=" * 42 + "\n")


def require_common():
    if "ensure_gspread" not in globals():
        raise RuntimeError("도구 셀을 실행해 공통 라이브러리를 먼저 로드해 주세요.")


def ensure_gspread():
    global gc, sh
    if "gc" not in globals():
        auth.authenticate_user()
        creds, _ = default()
        gc = gspread.authorize(creds)
    if "sh" not in globals():
        sh = gc.open_by_url(SPREADSHEET_URL)
    return gc, sh


def init_session():
    require_common()
    return ensure_gspread()


def _extract_full_text(document: dict) -> str:
    return "".join(
        el["textRun"]["content"]
        for block in document["body"]["content"]
        if "paragraph" in block
        for el in block["paragraph"]["elements"]
        if "textRun" in el
    )


def _get_docs_service():
    global _docs_service, _creds
    if _creds is None:
        _creds, _ = default()
    if _docs_service is None:
        _docs_service = build("docs", "v1", credentials=_creds, cache_discovery=False)
    return _docs_service


def _parse_log_no_cell(val) -> int | None:
    try:
        return int(float(str(val).strip()))
    except (ValueError, TypeError):
        return None


def get_or_create_log_index_sheet():
    return get_or_create_worksheet(sh, APPLY_LOG_SHEET_NAME, APPLY_LOG_HEADERS)


def allocate_next_log_no(ws=None) -> int:
    """`반영로그` 시트 A열 기준 다음 No."""
    if ws is None:
        ws = get_or_create_log_index_sheet()
    col = ws.col_values(1)
    nums = [_parse_log_no_cell(v) for v in col[1:]]
    nums = [n for n in nums if n is not None]
    return max(nums) + 1 if nums else 1


def encode_store_snapshot(rows: list) -> str:
    """상점현황 전체를 JSON 문자열로 (반영 직후 스냅샷)."""
    snapshot = [
        [r[0].strip(), r[1] if len(r) > 1 else ""]
        for r in rows
        if r and str(r[0]).strip()
    ]
    return json.dumps(snapshot, ensure_ascii=False)


def decode_store_snapshot(raw: str) -> list[list[str]]:
    data = json.loads(raw.strip())
    if not isinstance(data, list):
        raise ValueError("상점현황_JSON 형식이 올바르지 않습니다.")
    return [[str(c) for c in row] for row in data]


def _log_undo_col_1based(row: list) -> int:
    """실행취소 열 (1-based). 신규 6열=6, 구형 5열=5."""
    return 6 if len(row) >= 6 else 5


def append_log_index_row(
    log_no: int,
    timestamp: str,
    title: str,
    content_dict: dict,
    store_snapshot_json: str = "",
    *,
    ws=None,
) -> str | None:
    """반영로그 시트에 한 행 추가. 실패 시 오류 메시지."""
    try:
        if ws is None:
            ws = get_or_create_log_index_sheet()
        ws.append_row(
            [
                log_no,
                timestamp,
                title,
                json.dumps(content_dict, ensure_ascii=False),
                store_snapshot_json,
                "",
            ],
            value_input_option="USER_ENTERED",
        )
        return None
    except Exception as e:
        return f"{type(e).__name__}: {e}"


def get_log_index_entry(target_no: int) -> dict | None:
    """반영로그 한 건: changes, store_snapshot_json, meta."""
    try:
        ws = get_or_create_log_index_sheet()
        for i, row in enumerate(ws.get_all_values()[1:], start=2):
            if len(row) < 4:
                continue
            if _parse_log_no_cell(row[0]) != target_no:
                continue
            changes = None
            try:
                changes = json.loads(row[3].strip())
            except (json.JSONDecodeError, Exception):
                pass
            store_snapshot_json = ""
            if len(row) >= 6:
                store_snapshot_json = row[4].strip()
            undo_col = _log_undo_col_1based(row)
            undone_raw = row[undo_col - 1].strip() if len(row) >= undo_col else ""
            return {
                "row": i,
                "no": target_no,
                "timestamp": row[1] if len(row) > 1 else "",
                "title": row[2] if len(row) > 2 else "",
                "changes": changes,
                "store_snapshot_json": store_snapshot_json,
                "undone": undone_raw.upper() in ("Y", "YES", "예", "실행취소") or bool(undone_raw),
                "undo_col": undo_col,
            }
    except Exception:
        pass
    return None


def get_log_index_meta(target_no: int) -> dict | None:
    entry = get_log_index_entry(target_no)
    if not entry:
        return None
    return {
        "row": entry["row"],
        "no": entry["no"],
        "timestamp": entry["timestamp"],
        "title": entry["title"],
        "undone": entry["undone"],
        "undo_col": entry.get("undo_col", 6),
    }


def get_log_from_index(target_no: int) -> dict | None:
    entry = get_log_index_entry(target_no)
    if not entry:
        return None
    changes = entry.get("changes")
    return changes if isinstance(changes, dict) else None


def is_log_already_undone(target_no: int) -> bool:
    meta = get_log_index_meta(target_no)
    return bool(meta and meta["undone"])


def mark_log_undone(target_no: int) -> str | None:
    """반영로그 시트에 실행취소 표시."""
    try:
        entry = get_log_index_entry(target_no)
        if entry is None:
            return None
        ws = get_or_create_log_index_sheet()
        ws.update_cell(entry["row"], entry.get("undo_col", 6), "Y")
        return None
    except Exception as e:
        return f"{type(e).__name__}: {e}"


def append_machine_log(
    title_text: str,
    content_dict: dict,
    timestamp: str,
    store_rows: list | None = None,
) -> tuple[int, str | None]:
    """반영로그 시트에 변경 JSON + 상점현황 스냅샷 저장. (No, 오류) 반환."""
    try:
        index_ws = get_or_create_log_index_sheet()
        next_no = allocate_next_log_no(index_ws)
        snapshot = encode_store_snapshot(store_rows) if store_rows else ""
        index_err = append_log_index_row(
            next_no,
            timestamp,
            title_text,
            content_dict,
            snapshot,
            ws=index_ws,
        )
        if index_err:
            return -1, index_err
        return next_no, None
    except Exception as e:
        return -1, f"{type(e).__name__}: {e}"


def append_readable_log_to_doc(
    doc_id: str, title_text: str, timestamp: str, results_log: list[str], log_no: int,
    details_sort_helper: dict | None = None,
) -> str | None:
    """상점 반영기록(Docs) 삽입. 실패 시 오류 메시지, 성공 시 None."""
    if log_no is None or log_no < 1:
        return "유효한 반영로그 번호가 없어 상점 반영기록을 건너뜁니다."
    try:
        service = _get_docs_service()
        log_header = f"\n[No. {log_no}] {title_text} | {timestamp}\n"
        emoji_re = r"^[\U00010000-\U0010ffff✅🛒⚠️↩️]+\s*"
        clean_lines = [f"  • {re.sub(emoji_re, '', line)}\n" for line in results_log]
        detail_lines = []
        if details_sort_helper:
            cat_order = ["🏆 순위 보너스", "⚠️ 벌점 차감", "📅 화요일 미참", "🚫 미참여", "✨ 노참사 보너스", "🛡️ 부계정 토벌"]
            detail_lines.append("\n[상벌점 상세 지급 내역]\n")
            for cat in cat_order:
                if cat not in details_sort_helper:
                    continue
                items = details_sort_helper[cat]
                sorted_items = sorted(items, key=lambda x: x[1]) if cat == "🚫 미참여" else sorted(items, key=lambda x: x[0], reverse=True)
                cat_clean = re.sub(r"^[\U00010000-\U0010ffff✅🛒⚠️🏆✨🛡️📅🚫↩️]+\s*", "", cat)
                detail_lines.append(f"  {cat_clean}: {', '.join(item[1] for item in sorted_items)}\n")
        log_body = "".join(clean_lines) + "".join(detail_lines)
        divider = "─" * 40 + "\n"
        full_text = log_header + log_body + divider
        h_end, b_end = 1 + len(log_header), 1 + len(full_text)
        service.documents().batchUpdate(documentId=doc_id, body={"requests": [
            {"insertText": {"location": {"index": 1}, "text": full_text}},
            {"updateTextStyle": {"range": {"startIndex": 1, "endIndex": h_end},
                "textStyle": {"bold": True, "foregroundColor": {"color": {"rgbColor": {"green": 0.6}}}},
                "fields": "bold,foregroundColor"}},
            {"updateTextStyle": {"range": {"startIndex": h_end, "endIndex": b_end},
                "textStyle": {"bold": False, "foregroundColor": {}}, "fields": "bold,foregroundColor"}},
        ]}).execute()
        return None
    except Exception as e:
        return f"{type(e).__name__}: {e}"


def parse_score_from_cell(raw_val: str) -> float:
    match = re.search(r"([\+\-\[\]0-9.]+)", raw_val)
    if not match:
        return 0.0
    clean = re.sub(r"[^0-9.+-]", "", match.group(1))
    try:
        return float(clean) if clean else 0.0
    except ValueError:
        return 0.0


def format_score(value: float) -> str:
    if value == int(value):
        return f"+{int(value)}" if value > 0 else str(int(value))
    return f"[{value:+}]"


format_point = format_score


def sort_key_by_score(row: list) -> float:
    return parse_score_from_cell(row[1]) if len(row) >= 2 else -999999.0


def parse_store_text(text: str) -> list[list[str]]:
    restored = re.sub(r"\s*ㆍ", "\nㆍ", text).strip()
    result = []
    for line in restored.split("\n"):
        name = line.replace("ㆍ", "").strip()
        if not name:
            continue
        if ":" in name:
            n, v = name.split(":", 1)
            result.append([n.strip(), v.strip()])
        else:
            result.append([name, ""])
    return result


def parse_kv_input(text: str) -> list[tuple[str, str]]:
    if not text.strip():
        return []
    out = []
    for item in re.split(r"[\s,]+", text.strip()):
        if ":" in item:
            sub, main = item.split(":", 1)
            out.append((main.strip(), sub.strip()))
    return out


def col_letter(zero_based_idx: int) -> str:
    return chr(65 + zero_based_idx)


def get_log_by_no(target_no: int) -> dict | None:
    """실행취소용 JSON. `반영로그` 시트에서만 조회."""
    return get_log_from_index(target_no)


def load_alias_maps(map_data=None):
    if map_data is None:
        map_data = sh.worksheet("별명/다계정").get_all_values()
    alias_to_real = {row[2].strip(): row[3].strip() for row in map_data[1:] if len(row) > 3 and row[2].strip()}
    real_to_sheet = {row[1].strip(): row[0].strip() for row in map_data[1:] if len(row) > 1 and row[1].strip()}
    real_to_alias = {
        row[1].strip(): row[0].strip()
        for row in map_data[1:]
        if len(row) > 1 and row[1].strip() and row[0].strip() and row[0].strip() != row[1].strip()
    }
    sub_to_main = alias_to_real.copy()
    return alias_to_real, real_to_sheet, real_to_alias, sub_to_main


def load_mapping_sheet():
    ws = sh.worksheet("별명/다계정")
    return ws, ws.get_all_values()


def get_not_received(map_data: list) -> set[str]:
    return {row[4].strip() for row in map_data[1:] if len(row) > 4 and row[4].strip()}


def resolve_member(input_name: str, alias_to_real: dict, real_to_sheet: dict) -> tuple[str, str]:
    real = alias_to_real.get(input_name, input_name)
    target = real_to_sheet.get(real, real)
    return target, real


def load_store_sheet(title: str = "상점현황"):
    ws = sh.worksheet(title)
    data = ws.get_all_values()
    rows = [list(r) for r in data if any(c.strip() for c in r)]
    name_idx = {row[0].strip(): i for i, row in enumerate(rows) if row}
    return ws, rows, name_idx


def save_store_sheet(ws, rows: list) -> str | None:
    """상점현황 시트 저장. 실패 시 오류 메시지, 성공 시 None."""
    try:
        rows.sort(key=sort_key_by_score, reverse=True)
        ws.clear()
        ws.update(rows, range_name="A1")
        return None
    except Exception as e:
        return f"{type(e).__name__}: {e}"


def apply_delta_to_row(row: list, delta: float) -> tuple[float, float]:
    raw = row[1] if len(row) > 1 else "0"
    match = re.search(r"([\+\-\[\]0-9.]+)", raw)
    suffix = raw.replace(match.group(1), "", 1) if match else ""
    cur = parse_score_from_cell(raw)
    new = round(cur + delta, 2)
    row[1] = format_score(new) + suffix
    return cur, new


def split_input_tokens(text: str) -> list[str]:
    return [t for t in re.split(r"[\s,]+", text.strip()) if t]


def write_dual_log(
    title: str,
    log_dict: dict,
    results_log: list[str],
    timestamp: str,
    details_sort_helper: dict | None = None,
    store_rows: list | None = None,
) -> ApplyResult:
    """반영로그 시트 + 상점 반영기록 기록. 상점 반영 여부는 포함하지 않는다."""
    result = ApplyResult(tool=title, timestamp=timestamp, member_count=len(log_dict))
    log_no, index_err = append_machine_log(title, log_dict, timestamp, store_rows=store_rows)
    if index_err:
        result.machine_log_error = index_err
    else:
        result.machine_log_ok = True
        result.machine_log_no = log_no
    if result.machine_log_ok:
        readable_err = append_readable_log_to_doc(
            READABLE_DOC_ID, title, timestamp, results_log, log_no,
            details_sort_helper=details_sort_helper,
        )
        if readable_err:
            result.readable_log_error = readable_err
        else:
            result.readable_log_ok = True
    else:
        result.readable_log_error = "반영로그 실패로 상점 반영기록을 건너뜀"
    return result


def restore_store_from_backup(log_no: int) -> ApplyResult:
    """반영로그 No.의 상점현황_JSON으로 상점현황 시트 전체 복원."""
    ts = datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")
    title = f"No. {log_no} 백업 복원"
    result = ApplyResult(tool=title, timestamp=ts)

    entry = get_log_index_entry(log_no)
    if not entry:
        result.sheet_error = f"No. {log_no} 반영로그를 찾을 수 없습니다."
        print_apply_summary(result)
        return result

    raw = entry.get("store_snapshot_json", "").strip()
    if not raw:
        result.sheet_error = "이 로그에 상점현황_JSON 백업이 없습니다. (구형 로그 또는 저장 실패)"
        print_apply_summary(result)
        return result

    try:
        rows = decode_store_snapshot(raw)
        ws_store, _, _ = load_store_sheet()
        member_count = len([r for r in rows if r and r[0].strip()])
        result.member_count = member_count
        print(f"\n[{title}] 상점현황 {member_count}명 복원 (No. {log_no} 시점)")
        sheet_err = save_store_sheet(ws_store, rows)
        if sheet_err:
            result.sheet_error = sheet_err
        else:
            result.sheet_ok = True
            result.warnings.append("백업 복원은 반영로그·상점 반영기록에 새 No.를 남기지 않습니다.")
    except Exception as e:
        result.sheet_error = f"{type(e).__name__}: {e}"

    print_apply_summary(result)
    return result


def undo_store_log(log_no: int) -> ApplyResult | None:
    """반영로그 No.의 변경 JSON을 역반영하여 실행취소."""
    log_data = get_log_from_index(log_no)
    if not log_data:
        print(f"[알림] No. {log_no} 변경 JSON을 `반영로그`에서 찾을 수 없습니다.")
        return None

    ts = datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")
    ctx = open_store_context()
    undo_log_dict = {}
    for member, point in log_data.items():
        if not isinstance(point, (int, float)):
            continue
        target, _ = resolve_member(member, ctx["alias_to_real"], ctx["real_to_sheet"])
        undo = -float(point)
        undo_log_dict[target] = undo
        if target not in ctx["name_idx"]:
            print(f"[알림] '{target}'님을 상점현황에서 찾을 수 없어 건너뜁니다.")
            continue
        row = ctx["rows"][ctx["name_idx"][target]]
        cur, new = apply_delta_to_row(row, undo)
        ctx["results_log"].append(f"↩️ {target}: {cur} → {new} ({point:+} 취소)")

    if not ctx["results_log"]:
        print("[알림] 실행취소할 항목이 없습니다.")
        return None

    ctx["log_for_doc"] = undo_log_dict
    result = store_finalize(ctx, f"No. {log_no} 실행취소", ts, banner=f"No. {log_no} 실행취소")
    if result and result.sheet_ok:
        mark_err = mark_log_undone(log_no)
        if mark_err:
            print(f"[경고] 반영로그 '실행취소' 표시 실패: {mark_err}")
    return result


def open_store_context():
    ws_mapping, map_data = load_mapping_sheet()
    alias_to_real, real_to_sheet, _, _ = load_alias_maps(map_data)
    ws_store, rows, name_idx = load_store_sheet()
    return {
        "ws_mapping": ws_mapping,
        "map_data": map_data,
        "alias_to_real": alias_to_real,
        "real_to_sheet": real_to_sheet,
        "ws_store": ws_store,
        "rows": rows,
        "name_idx": name_idx,
        "results_log": [],
        "log_for_doc": {},
        "cancelled": False,
    }


def store_apply_delta(
    ctx: dict,
    input_name: str,
    delta: float,
    log_detail: str,
    *,
    allow_cancel: bool = True,
    skip_if_missing: bool = False,
    new_row_formatter=None,
) -> str | None:
    """상점현황에 delta만큼 반영. target_name 반환, 취소 시 None."""
    if ctx["cancelled"]:
        return None
    target, _ = resolve_member(input_name, ctx["alias_to_real"], ctx["real_to_sheet"])
    ctx["log_for_doc"][target] = ctx["log_for_doc"].get(target, 0) + delta

    if target in ctx["name_idx"]:
        row = ctx["rows"][ctx["name_idx"][target]]
        cur, new = apply_delta_to_row(row, delta)
        ctx["results_log"].append(f"✅ {target}: {cur} → {new} ({log_detail})")
        return target

    if skip_if_missing:
        print(f"[알림] '{target}'님을 상점현황에서 찾을 수 없어 건너뜁니다.")
        return target

    opts = "(예/아니요/취소)" if allow_cancel else "(예/아니요)"
    ans = input(f"'({target})' 상점목록에 없습니다. 추가할까요? {opts}: ").strip()
    if ans == "예":
        val = new_row_formatter(delta) if new_row_formatter else format_score(delta)
        ctx["rows"].append([target, val])
        ctx["name_idx"][target] = len(ctx["rows"]) - 1
        ctx["results_log"].append(f"✅ {target}: (신규) → {val}")
    elif ans == "취소" and allow_cancel:
        ctx["cancelled"] = True
    return target


def store_finalize(ctx: dict, title: str, timestamp: str, *, banner: str | None = None) -> ApplyResult | None:
    if ctx["cancelled"] or not ctx["results_log"]:
        return None
    display = banner or title
    result = ApplyResult(tool=display, timestamp=timestamp, member_count=len(ctx["log_for_doc"]))

    print(f"\n[{display}] 변경 내역 ({len(ctx['results_log'])}건)")
    for line in ctx["results_log"]:
        print(line)

    sheet_err = save_store_sheet(ctx["ws_store"], ctx["rows"])
    if sheet_err:
        result.sheet_error = sheet_err
    else:
        result.sheet_ok = True

    if result.sheet_ok:
        log_result = write_dual_log(
            title, ctx["log_for_doc"], ctx["results_log"], timestamp,
            store_rows=ctx["rows"],
        )
        result.merge_log(log_result)
    else:
        result.warnings.append("시트 저장 실패로 로그 기록을 건너뜀")

    print_apply_summary(result)
    return result


def apply_calc_points_to_store(
    calc_points: dict,
    map_data: list,
    ws_mapping,
    timestamp: str,
    season_label: str,
    details_sort_helper: dict | None = None,
) -> ApplyResult | None:
    """메인 정산: 계산된 상벌점을 상점현황에 일괄 반영."""
    alias_to_real, real_to_sheet, _, _ = load_alias_maps(map_data)
    not_received = get_not_received(map_data)
    ws_store, rows, name_idx = load_store_sheet()
    log_for_doc: dict[str, float] = {}
    results_log: list[str] = []
    cancelled = False

    for guild_member, point in calc_points.items():
        target, real = resolve_member(guild_member, alias_to_real, real_to_sheet)
        display = f"{real}({target})" if target != real else real
        if target in not_received:
            continue
        log_for_doc[target] = log_for_doc.get(target, 0) + point

        if target in name_idx:
            row = rows[name_idx[target]]
            cur, new = apply_delta_to_row(row, point)
            results_log.append(f"✅ {display}: {cur} → {new} ({point:+})")
        else:
            while True:
                ans = input(f"'({target})'은 상점 목록에 없습니다. 추가할까요? (예/아니요/취소): ").strip()
                if ans == "예":
                    rows.append([target, format_point(point)])
                    results_log.append(f"✅ {display}: (신규) → {format_point(point)}")
                    break
                if ans == "아니요":
                    nr_col = ws_mapping.col_values(5)
                    ws_mapping.update(values=[[target]], range_name=f"E{len(nr_col) + 1}")
                    break
                if ans == "취소":
                    cancelled = True
                    break
                print("[알림] '예', '아니요' 또는 '취소'로 입력해주세요.")
            if cancelled:
                break

    if cancelled:
        return None

    log_title = f"{season_label} 시즌 상벌점 반영"
    result = ApplyResult(
        tool=log_title,
        timestamp=timestamp,
        member_count=len(log_for_doc),
    )

    print(f"\n[{log_title}] 변경 내역 ({len(results_log)}건)")
    for line in results_log:
        print(line)

    sheet_err = save_store_sheet(ws_store, rows)
    if sheet_err:
        result.sheet_error = sheet_err
    else:
        result.sheet_ok = True

    if result.sheet_ok:
        log_result = write_dual_log(
            log_title,
            log_for_doc,
            results_log,
            timestamp,
            details_sort_helper=details_sort_helper,
            store_rows=rows,
        )
        result.merge_log(log_result)
    else:
        result.warnings.append("시트 저장 실패로 로그 기록을 건너뜀")

    print_apply_summary(result)
    return result


def get_or_create_worksheet(spreadsheet, title: str, headers: list | None = None):
    try:
        return spreadsheet.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        ncol = max(len(headers), 6) if headers else 6
        ws = spreadsheet.add_worksheet(title=title, rows="1000", cols=str(ncol))
        if headers:
            end = col_letter(len(headers) - 1)
            ws.update(range_name=f"A1:{end}1", values=[headers])
        return ws


def sync_kv(ws, input_text: str, sub_col: int, main_col: int) -> None:
    pairs = parse_kv_input(input_text)
    if not pairs:
        return
    current = ws.get_all_values()
    key_map = {
        row[main_col]: i + 1
        for i, row in enumerate(current)
        if len(row) > main_col and row[main_col].strip()
    }
    for main, sub in pairs:
        if main in key_map:
            ws.update(range_name=f"{col_letter(sub_col)}{key_map[main]}", values=[[sub]])
        else:
            next_row = len(ws.col_values(main_col + 1)) + 1
            left_col = col_letter(min(sub_col, main_col))
            right_col = col_letter(max(sub_col, main_col))
            row_vals = [sub, main] if sub_col < main_col else [main, sub]
            ws.update(range_name=f"{left_col}{next_row}:{right_col}{next_row}", values=[row_vals])
            key_map[main] = next_row


def sync_not_received(ws, input_text: str) -> None:
    names = [n.strip() for n in re.split(r"[\s,]+", input_text.strip()) if n.strip()]
    if not names:
        return
    existing = set(ws.col_values(5)[1:])
    to_add = [[n] for n in names if n not in existing]
    if to_add:
        next_row = len(ws.col_values(5)) + 1
        ws.update(range_name=f"E{next_row}", values=to_add)


# ---------------------------------------------------------------------------
# 공통 규칙 (`규칙` 시트 2행 단일) · 설정 (`설정` 시트)
# ---------------------------------------------------------------------------

RULES_SHEET_NAME = "규칙"
SETTINGS_SHEET_NAME = "설정"

# `규칙` 시트 1행 헤더 · 2행에 규칙 값 1세트만 둡니다.
RULES_SHEET_HEADERS = [
    "용_rate", "용_allow", "기_rate", "기_allow", "감_rate", "감_allow",
    "화요일_회당", "노참1", "노참2+", "전체1등", "전체2등", "전체3등",
    "보스1등", "보스2등", "3관왕", "부캐미참", "본캐미참",
]


def _rules_col_index(header: list[str], name: str) -> int:
    name_clean = name.strip().replace(" ", "")
    for i, h in enumerate(header):
        if h.strip().replace(" ", "") == name_clean:
            return i
    return -1


def _parse_num(raw: str, as_float: bool = False):
    raw = str(raw).strip()
    if not raw:
        return None
    clean = raw.replace(",", "")
    return float(clean) if as_float else int(float(clean))


def _rules_cell(row: list, header: list[str], col_name: str, *, as_float: bool = False):
    idx = _rules_col_index(header, col_name)
    if idx < 0 or idx >= len(row):
        return None
    return _parse_num(row[idx], as_float=as_float)


def _row_to_rules_dict(header: list[str], row: list[str]) -> dict:
    def need(col_name: str, **kwargs):
        v = _rules_cell(row, header, col_name, **kwargs)
        if v is None:
            raise ValueError(f"규칙 시트 2행에 '{col_name}' 값이 없습니다.")
        return v

    return {
        "bosses": {
            "dragon": {"rate": need("용_rate"), "allow": need("용_allow")},
            "machine": {"rate": need("기_rate"), "allow": need("기_allow")},
            "abyss": {"rate": need("감_rate"), "allow": need("감_allow")},
        },
        "tuesday_per_count": need("화요일_회당", as_float=True),
        "no_acc_1": need("노참1"),
        "no_acc_2plus": need("노참2+"),
        "overall_rank": {1: need("전체1등"), 2: need("전체2등"), 3: need("전체3등")},
        "boss_rank_1": need("보스1등"),
        "boss_rank_2": need("보스2등"),
        "triple_crown": need("3관왕"),
        "sub_absent_penalty": need("부캐미참"),
        "main_absent_penalty": need("본캐미참"),
    }


def load_rules(settlement_season: str | None = None) -> dict:
    """
    `규칙` 시트 1행=헤더, 2행=공통 규칙 1세트.
    settlement_season은 스냅샷·미리보기용(정산 시즌 ID)이며 규칙 조회와 무관.
    """
    ws = sh.worksheet(RULES_SHEET_NAME)
    rows = [r for r in ws.get_all_values() if any(c.strip() for c in r)]
    if len(rows) < 2:
        raise ValueError(
            f"'{RULES_SHEET_NAME}' 시트에 헤더(1행)와 규칙 데이터(2행)가 필요합니다.\n"
            f"헤더 예시: {', '.join(RULES_SHEET_HEADERS)}"
        )
    header = rows[0]
    data_row = next((row for row in rows[1:] if any(c.strip() for c in row)), None)
    if data_row is None:
        raise ValueError(f"'{RULES_SHEET_NAME}' 시트 2행에 규칙 값을 입력해 주세요.")
    rules = _row_to_rules_dict(header, data_row)
    if settlement_season and settlement_season.strip():
        rules["settlement_season"] = settlement_season.strip()
    return rules


def load_rules_for_season(season_id: str) -> dict:
    """하위 호환. 규칙은 시트 2행 단일, season_id는 정산 시즌(스냅샷)용."""
    return load_rules(settlement_season=season_id)


def rules_to_legacy_rules(rules: dict) -> dict[str, list[int]]:
    """메인 정산 process_boss_data 호환용 {boss_id: [rate, allow]}."""
    return {
        boss_id: [rules["bosses"][boss_id]["rate"], rules["bosses"][boss_id]["allow"]]
        for boss_id in ("dragon", "machine", "abyss")
    }


def print_rules_preview(rules: dict) -> None:
    b = rules["bosses"]
    print("\n" + "=" * 42)
    if rules.get("settlement_season"):
        print(f"[적용 규칙] 공통 (정산 시즌 {rules['settlement_season']})")
    else:
        print("[적용 규칙] 공통")
    print("=" * 42)
    print(f"  용(드래곤): rate {b['dragon']['rate']}%, allow {b['dragon']['allow']}")
    print(f"  기(기계신): rate {b['machine']['rate']}%, allow {b['machine']['allow']}")
    print(f"  감(감초):   rate {b['abyss']['rate']}%, allow {b['abyss']['allow']}")
    print(f"  화요일 미참: {rules['tuesday_per_count']}점/회")
    print(f"  노참사 보너스: 1보스 +{rules['no_acc_1']}, 2보스+ +{rules['no_acc_2plus']}")
    r = rules["overall_rank"]
    print(f"  전체 순위: 1등 +{r[1]}, 2등 +{r[2]}, 3등 +{r[3]}")
    print(f"  보스 순위: 1등 +{rules['boss_rank_1']}, 2등 +{rules['boss_rank_2']}")
    print(f"  3관왕: +{rules['triple_crown']}")
    print(f"  미참 벌점: 부캐 {rules['sub_absent_penalty']}점/회, 본캐 {rules['main_absent_penalty']}점/회")
    print("=" * 42 + "\n")


def rules_snapshot_rows(rules: dict) -> list[list[str]]:
    """시즌 시트 상단에 붙일 규칙 스냅샷 행."""
    rows = [["=== 적용 규칙 (정산 시점 스냅샷) ===", ""]]
    if rules.get("settlement_season"):
        rows.append(["정산 시즌", rules["settlement_season"]])
    rows.append(["JSON", json.dumps(rules, ensure_ascii=False)])
    rows.append([""])
    return rows


def load_settings() -> dict:
    """
    `설정` 시트 A열=키, B열=값. 없으면 기본값.
    """
    defaults = {"min_rows_per_boss": 30}
    try:
        ws = sh.worksheet(SETTINGS_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        return defaults.copy()
    for row in ws.get_all_values()[1:]:
        if len(row) < 2 or not row[0].strip():
            continue
        key, val = row[0].strip(), row[1].strip()
        if key in ("시트_최소행", "min_rows_per_boss") and val:
            defaults["min_rows_per_boss"] = int(float(val))
    return defaults


def print_settings_sheet_template() -> None:
    print("[설정 시트 예시] A1:B2")
    print("키\t값")
    print("시트_최소행\t30")


