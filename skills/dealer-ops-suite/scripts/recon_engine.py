import openpyxl, re, os, sys
from collections import defaultdict, Counter
from datetime import date as _date, datetime as _datetime, timedelta as _timedelta
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from decimal import Decimal
from settlement_rules import RateBook, RuleViolation, calculate_consignment, jsonable, money, normalize_business_type

# ==================== 可配置项（通过环境变量 RECON_CFG 传入 JSON）====================
CFG = {}
_CFG_ERR = ""
try:
    _CFG_RAW = os.environ.get("RECON_CFG", "") or ""
    if _CFG_RAW.strip():
        _parsed = json.loads(_CFG_RAW)
        if not isinstance(_parsed, dict):
            _CFG_ERR = "顶层必须是 JSON 对象"
        else:
            CFG = _parsed
except Exception as _e:
    _CFG_ERR = str(_e)

_SK = os.path.dirname(os.path.abspath(__file__))

# ---- v1.7.0：业务口径由 _口径确认.json 经 load_master 注入，规程不再写死 ----
KJ = CFG.get("koujing") or {}
V17 = bool((CFG.get("v17") or {}).get("enabled"))
KJD = KJ.get("decisions") or {}
EXPIRY_ON = bool((KJD.get("expiry") or {}).get("enabled"))
SHEET_MAP = KJD.get("sheet_map") or {"08": "试用装与中小样", "09": "效期预警"}
USAGE = KJD.get("usage") or {}
LEDGER = USAGE.get("ledger") or []


def _ledger_sum(kind, key):
    return sum(int(x.get(key, 0)) for x in LEDGER if x.get("kind") == kind)


def _ledger_closing(kind):
    return (_ledger_sum(kind, "opening") + _ledger_sum(kind, "shipped")
            - _ledger_sum(kind, "used"))


def _ledger_kinds():
    out = []
    for _x in LEDGER:
        _k = _x.get("kind")
        if _k and _k not in out:
            out.append(_k)
    return out


def _kinds_phrase(key):
    """按台账实际出现的物料类型生成「试用装 20 件、中小样 240 件、赠品 1 件」。"""
    return "、".join("%s %s 件" % (k, format(_ledger_sum(k, key) if key != "closing"
                                            else _ledger_closing(k), ","))
                    for k in _ledger_kinds())


def _policy_clause():
    _note = str((KJD.get("promo") or {}).get("policy_note") or "").strip()
    return "该部分货品" + _note + "，不参与货款结算" if _note else "该部分不参与货款结算"


def _cfg_fail(detail, missing=()):
    """缺配置时中止：stdout 必须为空，全部信息走 stderr。"""
    msg = [
        f"[中止] {detail}，未生成任何文件。",
        "生成方式（整行复制，先改 <> 里的内容）：",
    ]
    if os.name == "nt":
        msg += [
            "  Windows（PowerShell）：",
            '    $env:RECON_CFG = & py -3 "' + _SK + '\\load_master.py" "<主档的完整路径>" --period 202610 --asof 2026-11-02',
            '    py -3 "' + _SK + '\\recon_engine.py"',
        ]
    else:
        msg += [
            "  macOS / Linux（bash 或 zsh）：",
            f'    export RECON_CFG=$(python3 "{_SK}/load_master.py" "<主档的完整路径>" --period <202610> --asof <2026-11-02>)',
            f'    python3 "{_SK}/recon_engine.py"',
        ]
    if os.environ.get("RECON_DEBUG") and missing:
        msg.append("[debug] 缺失：" + "、".join(missing))
    sys.stderr.write("\n".join(msg) + "\n")
    raise SystemExit(1)


if _CFG_ERR:
    _cfg_fail("RECON_CFG 不是合法的配置 JSON：" + _CFG_ERR, ["RECON_CFG"])
if V17 and not KJ:
    _cfg_fail("缺少业务口径确认记录，请先完成阶段 G 问卡并写入 _口径确认.json",
              ["_口径确认.json"])

_need = [("party_a", "甲方主体"), ("party_b", "乙方主体"), ("period", "对账期间"),
         ("base_dir", "月份目录"), ("sources", "源文件位置")]
_miss = [lab for k, lab in _need if not CFG.get(k)]
if CFG.get("party_a") and not (CFG["party_a"] or {}).get("name"):
    _miss.append("甲方名称")
if CFG.get("party_b") and not (CFG["party_b"] or {}).get("name"):
    _miss.append("乙方名称")
if not CFG or _miss:
    _cfg_fail("本次对账缺少必要信息：" + (" / ".join(_miss) if _miss else "甲方主体 / 乙方主体 / 对账期间 / 源文件位置"), _miss)


_REAL_OUT = sys.stdout
sys.stdout = sys.stderr          # 过程中的一切输出走 stderr；失败时 stdout 保持 0 字节


def emit(msg=""):
    _REAL_OUT.write(str(msg) + "\n")
    _REAL_OUT.flush()


_PA = CFG.get("party_a", {})
_PB = CFG.get("party_b", {})
_CT = CFG.get("contract", {})
_PD = CFG.get("period", {})
_BK = CFG.get("bank", {})

from datetime import date as _date

PA_NAME = _PA.get("name") or ""
PA_TAX  = _PA.get("tax") or ""
PB_NAME = _PB.get("name") or ""
PB_TAX  = _PB.get("tax") or ""
CONTRACT_NO = _CT.get("no") or ""
STATEMENT_NO = str(CFG.get("statement_no") or "").strip()

# 期间三项必填；格式化文案与派生量一律由日期推导，不留业务字面量
P_START = _PD.get("start") or ""
P_END   = _PD.get("end") or ""
ASOF    = _PD.get("asof") or ""

_req2 = []
for _lab, _v in (("甲方纳税人识别号", PA_TAX), ("乙方纳税人识别号", PB_TAX),
                 ("合同编号", CONTRACT_NO), ("期间起始", P_START),
                 ("期间截止", P_END), ("数据基准日", ASOF)):
    if not _v:
        _req2.append(_lab)
_BK = _BK if isinstance(_BK, dict) else {}
if not (_BK.get("account_no") and _BK.get("bank_name")):
    _req2.append("甲方收款账户（开户行与账号）")
for _k, _lab in (("jia", "甲方挂账文件"), ("sales", "销售流水文件"), ("inventory", "库存文件")):
    if not (CFG.get("sources") or {}).get(_k):
        _req2.append(_lab)
_MASTER = CFG.get("master_path") or ""
if not _MASTER:
    _req2.append("主档路径（master_path）")
elif not os.path.exists(_MASTER):
    _req2.append("主档文件不存在：" + str(_MASTER))
if _req2:
    _cfg_fail("本次对账缺少必要信息：" + " / ".join(_req2), _req2)

try:
    _d_end  = _date.fromisoformat(str(P_END)[:10])
    _d_asof = _date.fromisoformat(str(ASOF)[:10])
    _d_start = _date.fromisoformat(str(P_START)[:10])
except ValueError:
    _cfg_fail("期间日期格式不正确（应为 YYYY-MM-DD）：start=%r end=%r asof=%r"
              % (P_START, P_END, ASOF), ["期间日期"])

P_CN    = _PD.get("period_cn") or f"{_d_end.year} 年 {_d_end.month} 月"
ASOF_CN = _PD.get("asof_cn") or f"{_d_asof.year} 年 {_d_asof.month:02d} 月 {_d_asof.day:02d} 日"
_d_next = _d_end + _timedelta(days=1)
_days_add = (_d_asof - _d_next).days + 1
ADD_BACK = _days_add > 0
NEXT_LABEL = (f"{_d_next.month} 月 {_d_next.day}—{_d_asof.day} 日" if ADD_BACK
              else f"{_d_next.month} 月 1—2 日")
NEXT_FULL = (f"{_d_next.year} 年 {_d_next.month} 月 {_d_next.day}—{_d_asof.day} 日"
             if ADD_BACK else f"{_d_next.month} 月 1—2 日")
PERIOD  = f"{_d_end.year:04d}{_d_end.month:02d}"

BK_NAME = _BK.get("account_name") or PA_NAME
BK_NO   = _BK.get("account_no") or ""
BK_BANK = _BK.get("bank_name") or ""

# ---- 合同条款：以主档「合同」表为准；缺失时不得静默套用演示值 ----
CT_METHOD = str(_CT.get("method") or "").strip()
CT_BASIS  = str(_CT.get("basis") or "").strip()
try:
    BUSINESS_TYPE = normalize_business_type(CFG.get("business_type") or CT_METHOD)
except RuleViolation as _e:
    _cfg_fail(f"业务路由阻断[{_e.code}] {_e.message}")
if BUSINESS_TYPE != "consignment":
    _cfg_fail("月结买断必须使用 monthly_recon_engine.py，不得套用铺货代销公式")
CT_METHOD_TXT = CT_METHOD or "（主档未登记，以合同为准）"
if CT_METHOD and CT_BASIS:
    SETTLE_TERMS_TXT = f"结算方式{CT_METHOD}、结算基数{CT_BASIS}"
elif CT_METHOD:
    SETTLE_TERMS_TXT = f"结算方式{CT_METHOD}（结算基数主档未登记，以合同为准）"
elif CT_BASIS:
    SETTLE_TERMS_TXT = f"结算基数{CT_BASIS}（结算方式主档未登记，以合同为准）"
else:
    SETTLE_TERMS_TXT = "结算方式与结算基数主档未登记，以合同为准"


def _term_workdays(value, what):
    """异议期/签回期：主档有正整数就用，否则显式写“合同约定的…”，不印固定演示值。"""
    try:
        n = int(float(str(value).strip()))
    except (TypeError, ValueError):
        n = 0
    return f"{n} 个工作日" if n > 0 else f"合同约定的{what}"


CT_SIGN_TXT = _term_workdays(_CT.get("signback_days"), "签回期限")
CT_OBJ_TXT  = _term_workdays(_CT.get("objection_days"), "异议期限")

SRC = CFG.get("sources", {})
F_JIA = SRC.get("jia") or ""
F_OCT = SRC.get("sales") or ""
F_INV = SRC.get("inventory") or ""

BASE = CFG["base_dir"]

# ---------------- 结构映射（可选，向后兼容）----------------
# load_master.py 在客户根目录发现 _数据映射.json 时会读入并放进 CFG["mapping"]。
# 无 mapping 时本段全部不生效，引擎行为与旧版完全一致。
_MAP = CFG.get("mapping") if isinstance(CFG.get("mapping"), dict) else None
if CFG.get("mapping") is not None and _MAP is None:
    _cfg_fail("RECON_CFG.mapping 必须是对象", ["mapping"])
_SMAP, _MMAP = {}, {}
if _MAP:
    if _MAP.get("schema") not in (None, "dealer-ops/data-mapping@1"):
        _cfg_fail("RECON_CFG.mapping.schema 不支持：%r" % (_MAP.get("schema"),), ["mapping.schema"])
    if isinstance(_MAP.get("sources"), dict):
        _SMAP = _MAP["sources"]
    if isinstance(_MAP.get("master"), dict):
        _MMAP = _MAP["master"]
    try:
        from schema_common import find_column as _map_find_column
    except Exception as _map_import_error:  # noqa: BLE001 - 依赖缺失时中文中止
        _cfg_fail("读取 _数据映射.json 需要 schema_common.py（%s）" % _map_import_error,
                  ["schema_common.py"])


def _map_bail(detail):
    """映射与真实文件不一致时中止：明确指出是哪一项，stdout 保持 0 字节。"""
    sys.stderr.write("\n[映射错误] " + detail + "\n"
                     "请核对 _数据映射.json（可先用 apply_mapping.py 做 dry-run 校验），"
                     "未生成任何文件。\n")
    raise SystemExit(1)


def _src_entry(role):
    entry = _SMAP.get(role)
    if isinstance(entry, dict) and entry:
        return entry
    if role != "next_month":
        return {}
    for parent in ("inventory", "jia"):
        parent_entry = _SMAP.get(parent)
        if isinstance(parent_entry, dict) and isinstance(parent_entry.get("next_month"), dict):
            return parent_entry["next_month"]
    return {}


def _src_label(role):
    """映射项在错误信息里的可定位路径（next_month 可能挂在 inventory/jia 下）。"""
    if role != "next_month":
        return "sources." + role
    for parent in ("inventory", "jia"):
        parent_entry = _SMAP.get(parent)
        if isinstance(parent_entry, dict) and isinstance(parent_entry.get("next_month"), dict):
            return "sources.%s.next_month" % parent
    return "sources.next_month"


def _src_file(role, default_name):
    """来源文件路径：映射 sources.<role>.file 优先，其次按关键词发现结果（旧行为）。"""
    entry = _src_entry(role)
    hint = entry.get("file")
    if not hint:
        return os.path.join(BASE, default_name) if default_name else ""
    hint = str(hint)
    if os.path.isabs(hint):
        candidates = [hint]
    else:
        candidates = [os.path.join(BASE, hint),
                      os.path.join(os.path.dirname(BASE), hint),
                      os.path.join(BASE, os.path.basename(hint)),
                      os.path.join(os.path.dirname(BASE), os.path.basename(hint))]
    for candidate in candidates:
        if os.path.exists(candidate):
            return candidate
    return candidates[0]


def _sheet_name_of(spec, names, role, path):
    label = _src_label(role)
    if isinstance(spec, int) and not isinstance(spec, bool):
        if not 1 <= spec <= len(names):
            _map_bail("%s.sheet 指定第 %r 张工作表，但文件 %s 只有 %d 张：%s"
                      % (label, spec, path, len(names), names))
        return names[spec - 1]
    name = str(spec)
    if name not in names:
        _map_bail("%s.sheet 指定的工作表「%s」在文件 %s 中不存在；实际工作表：%s"
                  % (label, name, path, names))
    return name


def _src_sheet(workbook, role, path, keys, exclude=()):
    entry = _src_entry(role)
    spec = entry.get("sheet")
    if spec in (None, ""):
        return _pick_sheet(workbook, keys, exclude)
    name = _sheet_name_of(spec, workbook.sheetnames, role, path)
    return workbook[name], name


def _check_header_row(header_row, role, max_rows):
    label = _src_label(role)
    if isinstance(header_row, bool) or not isinstance(header_row, int) or header_row < 1:
        _map_bail("%s.header_row 必须是大于等于 1 的整数（Excel 行号），当前 %r"
                  % (label, header_row))
    if header_row > max_rows:
        _map_bail("%s.header_row=%d 超出源表行数 %d" % (label, header_row, max_rows))
    return header_row


def _src_header(worksheet, role, path, keys, max_scan=6):
    entry = _src_entry(role)
    header_row = entry.get("header_row")
    if header_row in (None, ""):
        return _find_header(worksheet, keys, max_scan)
    header_row = _check_header_row(header_row, role, worksheet.max_row)
    values = [worksheet.cell(row=header_row, column=column).value
              for column in range(1, worksheet.max_column + 1)]
    return header_row, values


def _src_header_rows(matrix, role, keys, max_scan=6):
    entry = _src_entry(role)
    header_row = entry.get("header_row")
    if header_row in (None, ""):
        return _find_header_rows(matrix, keys, max_scan)
    header_row = _check_header_row(header_row, role, len(matrix))
    return header_row - 1, list(matrix[header_row - 1])


def _src_col(role, field, headers, keys, default, label="", deny=()):
    entry = _src_entry(role)
    columns = entry.get("columns")
    if not isinstance(columns, dict) or field not in columns:
        return _col(headers, keys, default, label, deny)
    index = _map_find_column(headers, columns[field])
    if index is None:
        actual = [str(value).strip() for value in headers
                  if value is not None and str(value).strip()]
        _map_bail("%s.columns.%s 指向的列 %r 在工作表表头中不存在；实际表头：%s"
                  % (_src_label(role), field, columns[field], actual))
    return index


_MASTER_SHEET_ALIAS = {}
if _MAP:
    _declared_sheets = _MMAP.get("sheets") if isinstance(_MMAP.get("sheets"), dict) else {}
    _MASTER_SHEET_ALIAS = {str(canonical): str(actual)
                           for actual, canonical in _declared_sheets.items()}


def _master_sheet(name):
    """主档工作表名：映射 master.sheets 声明了改名时用实际名，否则用规范名（旧行为）。"""
    return _MASTER_SHEET_ALIAS.get(name, name) if _MAP else name


_P_JIA = _src_file("jia", F_JIA)
_P_OCT = _src_file("sales", F_OCT)
_P_INV = _src_file("inventory", F_INV)

OUT = CFG.get("out_file") or os.path.join(
    BASE, _PD.get("out_name") or f"{_d_end.year}{_d_end.month:02d}铺货代销对账.xlsx")

_OUT_EXISTED = os.path.exists(OUT)
_OUT_LABEL = "[覆盖]" if _OUT_EXISTED else "[新建]"
if _OUT_EXISTED and not CFG.get("overwrite"):
    _stem, _ext = os.path.splitext(OUT)
    _n = 2
    while os.path.exists(f"{_stem}_v{_n}{_ext}"):
        _n += 1
    _new = f"{_stem}_v{_n}{_ext}"
    print(f"[保护] 目标文件已存在，本次写入 {os.path.basename(_new)}")
    print(f'       如需覆盖原文件，在 RECON_CFG 中加入 "overwrite": true')
    _OUT_LABEL = f"[覆盖→{os.path.basename(_new)}]"
    OUT = _new


NAVY, BLUE = "1F3864", "2F5597"
LBLUE, VLBLUE = "D9E2F3", "EDF2F9"
GREY, DGREY = "F2F2F2", "595959"
RED, LRED = "C00000", "FCE4E4"
GREEN, LGREEN = "2E6B30", "E8F1E8"
AMBER, LAMBER = "BF8F00", "FFF8E1"
WHITE, BORDC = "FFFFFF", "BFBFBF"
FN = "微软雅黑"

THIN = Side(style="thin", color=BORDC)
MED  = Side(style="medium", color="404040")

def F(sz=10, b=False, c="000000", it=False):
    return Font(name=FN, size=sz, bold=b, color=c, italic=it)

def fill(c): return PatternFill("solid", fgColor=c)

def A(h="center", v="center", wrap=True, ind=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=ind)

def box(ws, r1, c1, r2, c2, side=THIN):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = Border(
                left=side if c == c1 else cell.border.left,
                right=side if c == c2 else cell.border.right,
                top=side if r == r1 else cell.border.top,
                bottom=side if r == r2 else cell.border.bottom)

def mput(ws, r1, c1, r2, c2, v, font=None, al=None, fl=None, border=THIN):
    if r1 != r2 or c1 != c2:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    cell = ws.cell(row=r1, column=c1, value=v)
    if font: cell.font = font
    if al:   cell.alignment = al
    if fl:
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).fill = fill(fl)
    if border: box(ws, r1, c1, r2, c2, border)
    return cell

def _fit_h(txt, cpl=30.0):
    """按显示宽度估算行高：中文=2、其他=1。"""
    _w = sum(2 if ord(ch) > 127 else 1 for ch in str(txt))
    import math as _m
    return max(20, min(88, int(_m.ceil(_w / (cpl * 2)) * 13 + 5)))


def rmb_upper(amount):
    """金额 → 人民币大写。舍入统一走 settlement_rules.money（HALF_UP），
    元/角/分用 Decimal 精确分解，不再经 float 减法（0.34 类尾差曾靠 round() 补偿）。"""
    m = money(amount)  # Decimal，两位小数，HALF_UP
    if abs(m) < Decimal("0.005"): return "零元整"
    prefix = "负" if m < 0 else ""
    m = abs(m)
    digits = "零壹贰叁肆伍陆柒捌玖"
    unit1 = ["", "拾", "佰", "仟"]
    secs = ["", "万", "亿", "万亿"]
    integer = int(m)
    cents = int(m * 100) % 100  # m 已量化到 0.01，×100 精确无误差
    jiao, fen = divmod(cents, 10)
    if integer == 0:
        yuan = "零"
    else:
        s = str(integer); n = len(s)
        fl_ = n % 4 or 4
        groups = [s[:fl_]]; i = fl_
        while i < n:
            groups.append(s[i:i + 4]); i += 4
        nsec = len(groups); out = []
        for gi, g in enumerate(groups):
            if set(g) == {"0"}: continue
            gs = ""; zero = False
            for j, ch in enumerate(g):
                d = int(ch); pos = len(g) - 1 - j
                if d == 0: zero = True
                else:
                    if zero and gs: gs += "零"
                    gs += digits[d] + unit1[pos]; zero = False
            if gi > 0 and int(g) < 1000 and out: gs = "零" + gs
            out.append(gs + secs[nsec - 1 - gi])
        yuan = "".join(out)
    res = yuan + "元"
    if jiao == 0 and fen == 0: res += "整"
    else:
        if jiao: res += digits[jiao] + "角"
        elif fen: res += "零"
        res += digits[fen] + "分" if fen else "整"
    return prefix + res

# ==================== 读取源数据 ====================
def norm(s):
    s = str(s).replace("（", "(").replace("）", ")").replace(" ", "").replace("\u3000", "")
    return re.sub(r"\(限定版\)$", "", s)


def _pick_sheet(wb, keys, exclude=()):
    """按关键词找工作表；命中多张时中止并列出候选，不做猜测（缺失即阻断）。"""
    hits = [w for w in wb.worksheets
            if any(k in w.title for k in keys) and not any(x in w.title for x in exclude)]
    if len(hits) == 1:
        return hits[0], hits[0].title
    if len(hits) > 1:
        raise SystemExit(f"[中止] 含关键词 {list(keys)} 的工作表有 {len(hits)} 张："
                         f"{[w.title for w in hits]}，无法确定用哪一张。"
                         f"请只保留一张（或改表名），不要依赖脚本猜测。")
    cand = [w for w in wb.worksheets if not any(x in w.title for x in exclude) and w.max_row > 1]
    if not cand:
        raise SystemExit(f"[中止] 未找到含 {keys} 的工作表，且无含数据的候选表。实际工作表：{wb.sheetnames}")
    if len(cand) > 1:
        raise SystemExit(f"[中止] 未找到含 {keys} 的工作表，但存在 {len(cand)} 张含数据表："
                         f"{[w.title for w in cand]}。请确认用哪张（可改表名含关键词）。")
    ws = cand[0]
    print(f"⚠ 未找到含 {keys} 的工作表，使用唯一含数据表「{ws.title}」")
    return ws, ws.title


def _find_header(ws, keys, max_scan=6):
    """在前几行里找表头行，返回 (行号, 表头列表)。"""
    for r in range(1, min(ws.max_row, max_scan) + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if any(v is not None and any(k in str(v) for k in keys) for v in vals):
            return r, vals
    return 1, [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]


def _col(headers, keys, default, label="", deny=()):
    """先匹配完整表头，再用长关键词；同级多列命中即阻断，不按列序猜测。"""
    def _header_text(value):
        text = re.sub(r"\s+", "", str(value or ""))
        return re.sub(r"[(（][^)）]*[)）]", "", text)

    candidates = []
    for i, header in enumerate(headers):
        if header is None:
            continue
        raw = str(header)
        if deny and any(d in raw for d in deny):
            continue
        text = _header_text(raw)
        if text:
            candidates.append((i, text))
    for key in keys:
        target = _header_text(key)
        hits = [i for i, header in candidates if header == target]
        if len(hits) > 1:
            _bail(f"数据表表头「{key}」命中多列，无法唯一取数。实际表头：{headers}")
        if hits:
            return hits[0]
    for key in keys:
        target = _header_text(key)
        if target in ("数量", "销量", "时间", "日期"):
            continue  # 泛化业务词只接受完整表头，防止命中“发货数量／发货时间”
        hits = [i for i, header in candidates if target in header]
        if len(hits) > 1:
            _bail(f"数据表中含「{key}」的列不唯一，已停止取数。实际表头：{headers}")
        if hits:
            return hits[0]
    if label:
        _bail(f"数据表未找到含「{'/'.join(keys)}」的列（{label}）。实际表头：{[str(h) for h in headers if h]}。"
              f"请核对导出格式，或在配置中调整源文件")
    return default


def _cell_of(row, idx):
    if idx is None or row is None or idx < 0 or idx >= len(row):
        return None
    return row[idx]


def _num_val(value, ctx):
    """源单元格 → 数字。必需字段空白即阻断；只有明确填写的数字 0 才是零。"""
    if value is None or (isinstance(value, str) and not value.strip()):
        _bail(f"{ctx} 为空。必需数量／金额缺失即阻断，不得按 0 猜算，请核对源表")
    if isinstance(value, bool):
        _bail(f"{ctx} 是布尔值（{value!r}），疑似表头/说明行混入数据区，请核对源表")
    if isinstance(value, (int, float)):
        f = float(value)
    else:
        s = str(value).strip().replace(",", "").replace("，", "")
        try:
            f = float(s)
        except ValueError:
            _bail(f"{ctx} 的值「{value!r}」不是数字。缺失即阻断，不做静默转换，请核对源表")
    if f != f or f in (float("inf"), float("-inf")):
        _bail(f"{ctx} 的值「{value!r}」不是有限数字，请核对源表")
    return f


def _qval(value, ctx):
    """源数量 → 保留实数（4 位小数）。数量不得 int 取整：静默取整会改结算金额。"""
    return round(_num_val(value, ctx), 4)


def _read_matrix(path, sheet=None, role="jia"):
    """把源文件读成二维列表：.xlsx/.xlsm 用 openpyxl，.xls 用 xlrd（按扩展名分派）。

    sheet 为映射显式指定的工作表名 / 序号（1 基）；为 None 时沿用旧行为（第一张工作表）。
    """
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        try:
            if sheet in (None, ""):
                ws = wb.worksheets[0]
            else:
                ws = wb[_sheet_name_of(sheet, wb.sheetnames, role, path)]
            return [list(r) for r in ws.iter_rows(values_only=True)]
        finally:
            wb.close()
    if ext == ".xls":
        try:
            import xlrd
        except ImportError:
            raise SystemExit("[缺少依赖] 读取 .xls 需要 xlrd（python3 -m pip install xlrd）；"
                             "或请甲方将挂账表另存为 .xlsx 后重试。")
        wb = xlrd.open_workbook(path)
        if sheet in (None, ""):
            sh = wb.sheet_by_index(0)
        else:
            sh = wb.sheet_by_name(_sheet_name_of(sheet, wb.sheet_names(), role, path))
        return [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    raise SystemExit(f"[中止] 不支持的源文件格式：{path}（仅支持 .xls / .xlsx / .xlsm）")


def _find_header_rows(rows, keys, max_scan=6):
    """在二维矩阵的前几行里找表头行（兼容导出文件首行是标题行）。"""
    for i, row in enumerate(rows[:max_scan]):
        if any(v is not None and any(k in str(v) for k in keys) for v in row):
            return i, list(row)
    return 0, list(rows[0]) if rows else []


def _iso_day(value):
    """销售时间 → YYYY-MM-DD；支持 datetime/date、ISO、斜杠、点号、yyyymmdd 字符串。"""
    if isinstance(value, _datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, _date):
        return value.isoformat()
    s = str(value or "").strip()
    if not s:
        return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d",
                "%Y.%m.%d", "%Y%m%d"):
        try:
            return _datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""


_PART = {}

try:
    from schema_common import normalize_barcode as _norm_bar
    from schema_common import locate_column as _locate_column
    from schema_common import ColumnAmbiguous as _ColumnAmbiguous
except Exception as _norm_import_error:  # noqa: BLE001 - 依赖缺失时中文中止
    raise SystemExit("[中止] 条码归一化／列定位需要 schema_common.py（%s）；请确认脚本目录完整"
                     % _norm_import_error)


def _bail(msg):
    """失败时先把已算出的关键数打到 stderr，便于定位。"""
    sys.stderr.write("\n[自检未通过] " + msg + "\n")
    if _PART:
        sys.stderr.write("已算出：期初 %s 件/%s 元；结算 %s 件/%s 元；差异 %s 件/%s 元\n"
                         % (_PART.get("q0"), _PART.get("open"), _PART.get("q10"),
                            _PART.get("settle"), _PART.get("diff"), _PART.get("diff_amt")))
    raise SystemExit(1)


def _require(path):
    if not os.path.exists(path):
        cands = sorted(f for f in os.listdir(BASE)
                       if f.lower().endswith((".xls", ".xlsx", ".xlsm"))) if os.path.isdir(BASE) else []
        raise SystemExit(
            f"\n[缺少源文件] {path}\n"
            f"目录 {BASE} 下现有 Excel：{cands or '（无）'}\n"
            f"请核对文件名，或在 RECON_CFG 的 sources 中指定。")

for _p in (_P_JIA, _P_OCT, _P_INV):
    _require(_p)

# ---------- 甲方挂账：只读取数量，单价唯一权威源是合同价格政策 ----------
_jia_mat = _read_matrix(_P_JIA, sheet=_src_entry("jia").get("sheet"), role="jia")
_hr0, _h0 = _src_header_rows(_jia_mat, "jia", ["商品名称", "未结算数量"])
c_name = _src_col("jia", "name", _h0, ["商品名称", "商品名", "名称", "品名", "商品"], 0,
                      label="甲方挂账·品名列", deny=("条码", "编码", "编号", "代码"))
c_qty = _src_col("jia", "qty", _h0, ["未结算数量", "数量", "现存量"], None, label="甲方挂账·数量列")
c_init = _src_col("jia", "init", _h0, ["初始数量", "初始"], c_qty)
# 新的甲方挂账表会带批次号 / 有效期 / 物料属性；缺列时按旧行为（正装、无批次）
c_lot = _src_col("jia", "lot", _h0, ["批次号", "批次", "批号"], None)
c_exp = _src_col("jia", "exp", _h0, ["有效期", "到期日", "失效日期"], None)
c_kind = _src_col("jia", "kind", _h0, ["物料属性", "物料类型", "品项类型", "商品类型"], None)
c_code = _src_col("jia", "code", _h0, ["物料编码", "客户物料号", "物料编号", "商品编码"], None)

_KIND_RULES = (("试用", "试用装"), ("小样", "中小样"), ("赠品", "配赠"),
               ("配赠", "配赠"), ("促销", "配赠"), ("陈列", "陈列"),
               ("正装", "正装"), ("正常", "正装"))
_KINDS = ("正装", "试用装", "中小样", "配赠", "陈列")
_PROMO_KINDS = ("试用装", "中小样", "配赠")


def _kind_of(value):
    """物料属性归一化；空值按正装（与旧版一致），未登记取值留给上层阻断。"""
    s_ = "" if value is None else str(value).strip()
    if not s_:
        return "正装"
    for key, name in _KIND_RULES:
        if key in s_:
            return name
    return s_


def _parse_day(value):
    """日期解析：datetime / date / Excel 序列号 / 常见字符串；识别不了返回 None。"""
    if value in (None, ""):
        return None
    if isinstance(value, _datetime):
        return value.date()
    if isinstance(value, _date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            return (_datetime(1899, 12, 30) + _timedelta(days=float(value))).date()
        except (OverflowError, ValueError):
            return None
    txt = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d",
                "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
        try:
            return _datetime.strptime(txt, fmt).date()
        except ValueError:
            continue
    return None


_LOT_SEP = re.compile(r"[/／|｜]")


def _split_lot(value):
    """批次字段常写成「批次号/有效期」合体（如 R0320B11/20290319）。

    返回 (批次号, 有效期或 None)：右段能解析成日期才拆，否则整串当批次号，
    避免把「A/2026」这类非日期批次误拆。
    """
    txt = "" if value is None else str(value).strip()
    if not txt:
        return "", None
    parts = [x.strip() for x in _LOT_SEP.split(txt) if x.strip()]
    if len(parts) >= 2:
        exp = _parse_day(parts[-1])
        if exp is not None:
            return "/".join(parts[:-1]), exp
    return txt, None


def _add_months(day, months):
    """生产日期 + 保质期(月) → 到期日；月末按当月最后一天收敛。"""
    y, m = divmod(day.month - 1 + int(months), 12)
    y += day.year
    m += 1
    last = [31, 29 if (y % 4 == 0 and (y % 100 != 0 or y % 400 == 0)) else 28,
            31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]
    return _date(y, m, min(day.day, last))

_jia_raw = []
for _ri, r in enumerate(_jia_mat[_hr0 + 1:], _hr0 + 2):
    _nm = _cell_of(r, c_name)
    nm = "" if _nm is None else str(_nm).strip()
    if not nm:
        continue
    if nm in ("合计", "小计", "总计", "合 计"):
        continue
    _lot, _lot_exp = _split_lot(_cell_of(r, c_lot) if c_lot is not None else None)
    _exp = (_parse_day(_cell_of(r, c_exp)) if c_exp is not None else None) or _lot_exp
    _jia_raw.append(dict(name=nm,
                         price=0.0,  # 仅占位；后续由合同价格表按日期赋值，不读甲方单价
                         amt=0.0,    # 甲方金额不参与结算或恒等式
                         qty=_qval(_cell_of(r, c_qty), f"甲方挂账第 {_ri} 行·未结算数量"),
                         init=_qval(_cell_of(r, c_init), f"甲方挂账第 {_ri} 行·初始数量"),
                         code="" if c_code is None or _cell_of(r, c_code) is None
                              else str(_cell_of(r, c_code)).strip(),
                         lot=_lot,
                         exp=_exp,
                         kind=_kind_of(_cell_of(r, c_kind) if c_kind is not None else None)))

# 物料属性三态：未登记取值一律阻断，不默认成正装或促销
_bad_kinds = sorted({x["kind"] for x in _jia_raw if x["kind"] not in _KINDS})
if _bad_kinds:
    _bail("甲方挂账表「物料属性」出现未登记取值：" + "、".join(_bad_kinds)
          + "。仅接受 正装／试用装／中小样／配赠／陈列（空值按正装），请补全源表后再对账")

# 同品名多批次（新挂账表按批次展开）属正常拆行，不做“重复挂账”告警
_HAS_LOT = any(x["lot"] for x in _jia_raw)
_jia_settle = [x for x in _jia_raw if x["kind"] not in _PROMO_KINDS]
_jia_promo = [x for x in _jia_raw if x["kind"] in _PROMO_KINDS]

# 同品名多行按数量合计，避免动销被双算
jia, _jm = [], {}
for _x in _jia_settle:
    _k = norm(_x["name"])
    if _k in _jm:
        _o = _jm[_k]
        if not _HAS_LOT:
            print(f"⚠ 挂账表中「{_x['name']}」出现多行，已按数量合计")
        _o["qty"] += _x["qty"]; _o["init"] += _x["init"]; _o["amt"] = float(money(_o["amt"] + _x["amt"]))
        _o["lots"].append((_x["lot"], _x["exp"], _x["qty"]))
    else:
        _jm[_k] = dict(_x, lots=[(_x["lot"], _x["exp"], _x["qty"])])
        jia.append(_jm[_k])

# 促销物料（试用装 / 中小样 / 配赠）：不参与货款结算与库存恒等式，另出附表 09
PROMO = {}
for _x in _jia_promo:
    _k = norm(_x["name"])
    if _k in PROMO:
        _p = PROMO[_k]
        _p["out_qty"] += _x["qty"]
        _p["amt"] = float(money(_p["amt"] + _x["amt"]))
        if _p["kind"] != _x["kind"]:
            print(f"⚠ 促销物料「{_x['name']}」存在多个物料属性（{_p['kind']} / {_x['kind']}），已按首次出现归类")
    else:
        PROMO[_k] = dict(name=_x["name"], kind=_x["kind"], out_qty=_x["qty"],
                         amt=_x["amt"], price=_x["price"], bar="")
    if _x["lot"] or _x["exp"]:
        PROMO[_k].setdefault("lots", []).append((_x["lot"], _x["exp"], _x["qty"]))

# 批次明细（仅正装）：促销物料是否入库不确定，不纳入效期预警
BATCH = []
for _x in _jia_settle:
    if (not _x["lot"] and not _x["exp"]) or _x["qty"] <= 0:
        continue
    BATCH.append(dict(name=_x["name"], bar="", code=_x.get("code", ""),
                      lot=_x["lot"] or "（未标批次）",
                      exp=_x["exp"], qty=_x["qty"], init=_x["init"], price=_x["price"]))
_BATCH_FROM_JIA = bool(BATCH)

# ---------- 乙方库存（表名与列均按关键词识别；有映射时按映射）----------
wb1 = openpyxl.load_workbook(_P_INV, data_only=True)
ws_inv, inv_sheet_name = _src_sheet(wb1, "inventory", _P_INV, ["库存"], exclude=("销售",))
_hr, _hinv = _src_header(ws_inv, "inventory", _P_INV, ["商品名称", "库存"])
k_name = _src_col("inventory", "name", _hinv, ["商品名称", "名称", "品名"], 0,
                  label="库存表·品名列")
k_bar = _src_col("inventory", "bar", _hinv, ["条码", "编码"], 1, label="库存表·条码列")
k_spec = _src_col("inventory", "spec", _hinv, ["规格"], 2)
k_cat = _src_col("inventory", "cat", _hinv, ["分类"], 3)
k_qty = _src_col("inventory", "qty", _hinv, ["库存", "数量"], 4, label="库存表·库存列")
k_retail = _src_col("inventory", "retail", _hinv, ["销售价", "零售价"], 5)
k_brand = _src_col("inventory", "brand", _hinv, ["品牌"], None)

inv = []
for _ri, r in enumerate(ws_inv.iter_rows(min_row=_hr + 1, values_only=True), _hr + 1):
    if not r or len(r) <= max(k_name, k_qty) or not r[k_name]:
        continue
    _bar = _norm_bar(_cell_of(r, k_bar))
    _retail = _cell_of(r, k_retail)
    _brand = _cell_of(r, k_brand)
    inv.append(dict(name=str(r[k_name]).strip(),
                    bar=_bar,
                    spec=_cell_of(r, k_spec) if _cell_of(r, k_spec) is not None else "",
                    cat=_cell_of(r, k_cat) if _cell_of(r, k_cat) is not None else "",
                    qty=_qval(_cell_of(r, k_qty), f"库存表第 {_ri} 行·库存数量"),
                    retail=(_num_val(_retail, f"库存表第 {_ri} 行·销售价")
                            if _retail is not None and str(_retail).strip() else 0.0),
                    brand=str(_brand).strip() if _brand is not None else ""))

# 先按条码聚合（同一商品被拆成两行、或换名不换码，都能合并）
_bydedup = {}
for _x in inv:
    _bk = _x["bar"] or ("NAME:" + norm(_x["name"]))
    if _bk in _bydedup:
        _bydedup[_bk]["qty"] += _x["qty"]
        if norm(_x["name"]) != norm(_bydedup[_bk]["name"]):
            print(f"⚠ 库存表内条码 {_x['bar']} 对应多个品名（{_bydedup[_bk]['name']} / {_x['name']}），已按条码合并")
        else:
            print(f"⚠ 库存表内「{_x['name']}」出现多行，已按数量合计")
    else:
        _bydedup[_bk] = dict(_x)
# 再按名称聚合（个别表无条码）
_dedup = {}
for _x in _bydedup.values():
    _k = norm(_x["name"])
    if _k in _dedup:
        _dedup[_k]["qty"] += _x["qty"]
    else:
        _dedup[_k] = _x
inv = list(_dedup.values())
inv_by_n = {norm(x["name"]): x for x in inv}
inv_by_bar = {x["bar"]: x for x in inv if x["bar"]}

# ---------- 乙方库存导出月次月销售（另一张表，表名随月变化）----------
nov = defaultdict(int)
ws_nov, nov_sheet_name = None, ""
_nov_skip = []
if not ADD_BACK:
    print(f"库存导出日 {ASOF} 不晚于期末 {P_END}，无需回加，跳过次月销售表")
else:
    ws_nov, nov_sheet_name = _src_sheet(wb1, "next_month", _P_INV, ["销售"], exclude=("库存",))
    _t = ws_nov.title or ""
    if str(_d_next.month) not in _t and "月" in _t:
        print(f"⚠ 回加工作表「{_t}」与推算区间 {_d_next.month} 月不一致，请核对库存导出日")
if ADD_BACK:
    _hr2, _hnov = _src_header(ws_nov, "next_month", _P_INV, ["商品名称", "销售数量"])
    n_name = _src_col("next_month", "name", _hnov, ["商品名称", "名称"], 0)
    n_bar = _src_col("next_month", "bar", _hnov, ["条码", "编码"], 1)
    n_num = _src_col("next_month", "qty", _hnov, ["销售数量", "数量", "销量"], 3,
                     label="次月销售表·销售数量列")
else:
    _hr2, _hnov, n_name, n_bar, n_num = 1, [], 0, 1, 3
nov = defaultdict(int)
for _ri, r in enumerate((ws_nov.iter_rows(min_row=_hr2 + 1, values_only=True) if ADD_BACK else []), _hr2 + 1):
    if not r or not _cell_of(r, n_name):
        continue
    _q = _num_val(_cell_of(r, n_num), f"次月销售表第 {_ri} 行·销售数量")
    if not _q:
        continue
    _bar = _norm_bar(_cell_of(r, n_bar))
    if not _bar:
        _iv = inv_by_n.get(norm(_cell_of(r, n_name)))
        _bar = _iv["bar"] if _iv else ""
    if not _bar:
        _nov_skip.append(str(_cell_of(r, n_name)).strip())
        continue
    nov[_bar] += round(_q, 4)
if _nov_skip:
    print(f"⚠ 次月销售表有 {len(_nov_skip)} 行无条码且品名无法匹配库存，已跳过（如："
          f"{'、'.join(_nov_skip[:5])}）；请核对是否影响回加")

# ---------- 乙方当期销售流水 ----------
wb2 = openpyxl.load_workbook(_P_OCT, data_only=True)
ws_oct, oct_sheet_name = _src_sheet(wb2, "sales", _P_OCT, [str(P_END)[5:7], "月"],
                                   exclude=("库存",))
_m_end = _date.fromisoformat(str(P_END)[:10]).month
_m_hit = re.search(r"(\d{1,2})\s*月", ws_oct.title or "")
if _m_hit and int(_m_hit.group(1)) != _m_end:
    _bail(f"销售工作表是「{ws_oct.title}」，但本次期间是 {_m_end} 月，疑似期间与源文件不匹配。"
          f"请把对应月份的销售数据放到本目录，或在配置中显式指定 sources.sales")

_hr3, _hoct = _src_header(ws_oct, "sales", _P_OCT, ["商品名称", "销售数量"])
o_id = _src_col("sales", "serial", _hoct, ["流水号", "单号"], 0)
o_time = _src_col("sales", "date", _hoct, ["销售时间", "时间", "日期"], 1)
o_store = _src_col("sales", "store", _hoct, ["门店", "店铺", "网点"], 2)
o_name = _src_col("sales", "name", _hoct, ["商品名称", "名称", "品名"], 3,
                  label="销售流水·品名列")
o_bar = _src_col("sales", "bar", _hoct, ["条码", "编码"], 4, label="销售流水·条码列")
o_spec = _src_col("sales", "spec", _hoct, ["规格"], 5)
o_unit = _src_col("sales", "unit", _hoct, ["单位"], 6)
o_num = _src_col("sales", "qty", _hoct, ["销售数量", "数量", "销量"], 7, label="销售流水·数量列")
o_brand = _src_col("sales", "brand", _hoct, ["品牌"], 8)
o_kind = _src_col("sales", "kind", _hoct, ["赠品", "试用", "物料属性", "物料类型"], None)
o_txn = _src_col("sales", "txn_type", _hoct, ["业务类型", "单据类型", "销售类型"], None)
o_orig = _src_col("sales", "original_sale_id", _hoct, ["原销售流水号", "原销售单号"], None)
o_orig_price = _src_col("sales", "original_unit_price", _hoct, ["原销售单价", "原结算价"], None)
o_orig_th = _src_col("sales", "original_decal_charged", _hoct, ["原销售已计提贴花", "原销售贴花已计提"], None)
o_orig_th_std = _src_col("sales", "original_decal_standard", _hoct, ["原销售贴花标准", "原贴花标准"], None)
_sales_use = defaultdict(int)
_AGGREGATE_LABELS = {"合计", "小计", "总计"}
_aggregate_rows_skipped = 0


def _optional_bool(value, label):
    if value in (None, ""):
        return None
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in ("1", "true", "是", "已计提", "有"):
        return True
    if text in ("0", "false", "否", "未计提", "无"):
        return False
    _bail(f"{label}只能填 是／否，收到：{value!r}")


detail = []
sales_rule_lines = []
oct_qty, oct_store = defaultdict(int), defaultdict(lambda: defaultdict(int))
oct_name, oct_spec, oct_unit = {}, {}, {}
oct_by_name = {}
for _ri, r in enumerate(ws_oct.iter_rows(min_row=_hr3 + 1, values_only=True), _hr3 + 1):
    if not r or not any(value is not None and str(value).strip() for value in r):
        continue
    # 与挂账、库存表一致：显式聚合行豁免并留痕；其余缺关键字段的业务行必须阻断。
    _cells = {norm(value) for value in r if value is not None and str(value).strip()}
    if any(value in _AGGREGATE_LABELS for value in _cells):
        _aggregate_rows_skipped += 1
        continue
    if _cell_of(r, o_id) is None or not str(_cell_of(r, o_id)).strip():
        _bail(f"销售流水第 {_ri} 行缺少流水号；有效业务行不得静默跳过，请补齐后重新对账")
    if _cell_of(r, o_name) is None or not str(_cell_of(r, o_name)).strip():
        _bail(f"销售流水第 {_ri} 行缺少商品名称；有效业务行不得静默跳过，请补齐后重新对账")
    raw_q = _qval(_cell_of(r, o_num), f"销售流水第 {_ri} 行·销售数量")
    txn_text = str(_cell_of(r, o_txn) or "销售").strip() if o_txn is not None else "销售"
    txn_kind = "return" if txn_text in ("退货", "冲销", "return") else "sale"
    if txn_kind == "sale" and txn_text not in ("销售", "动销", "sale"):
        _bail(f"销售流水第 {_ri} 行业务类型无法识别：{txn_text!r}")
    q = -abs(raw_q) if txn_kind == "return" else raw_q
    _bar = _norm_bar(_cell_of(r, o_bar))
    if not _bar:
        _bail(f"销售流水第 {_ri} 行缺少商品条码。销售流水必须带条码，不得按品名猜算，请核对导出文件")
    _store = _cell_of(r, o_store)
    _store = str(_store).strip() if _store not in (None, "") else "未标门店"
    _spec = _cell_of(r, o_spec)
    _unit = _cell_of(r, o_unit)
    _brand = _cell_of(r, o_brand)
    _original_id = str(_cell_of(r, o_orig) or "").strip() if o_orig is not None else ""
    _rule_line = {
        "id": str(_cell_of(r, o_id)).strip(), "barcode": _bar, "date": _cell_of(r, o_time),
        "quantity": abs(raw_q) if txn_kind == "return" else raw_q, "kind": txn_kind,
        "original_sale_id": _original_id,
        "original_decal_charged": _optional_bool(_cell_of(r, o_orig_th) if o_orig_th is not None else None,
                                                f"销售流水第 {_ri} 行·原销售已计提贴花"),
    }
    if o_orig_price is not None and _cell_of(r, o_orig_price) not in (None, ""):
        _rule_line["original_unit_price"] = _num_val(_cell_of(r, o_orig_price), f"销售流水第 {_ri} 行·原销售单价")
    if o_orig_th_std is not None and _cell_of(r, o_orig_th_std) not in (None, ""):
        _rule_line["original_decal_standard"] = _num_val(_cell_of(r, o_orig_th_std), f"销售流水第 {_ri} 行·原销售贴花标准")
    sales_rule_lines.append(_rule_line)
    detail.append([_cell_of(r, o_id), _cell_of(r, o_time), _store, _cell_of(r, o_name),
                   _bar, _spec if _spec is not None else "",
                   _unit if _unit is not None else "", q,
                   _brand if _brand is not None else "", txn_text, _original_id, None, None, ""])
    oct_qty[_bar] += q
    oct_store[_bar][_store] += q
    oct_name[_bar] = str(_cell_of(r, o_name)).strip()
    oct_spec[_bar] = _spec if _spec is not None else ""
    oct_unit[_bar] = _unit if _unit is not None else ""
    oct_by_name[norm(_cell_of(r, o_name))] = _bar
    if o_kind is not None:
        _kv = _cell_of(r, o_kind)
        if _kv and any(_w in str(_kv) for _w in ("赠品", "试用", "小样", "配赠", "促销")):
            _sales_use[_bar] += q
if _aggregate_rows_skipped:
    print(f"⚠ 销售流水豁免 {_aggregate_rows_skipped} 行合计／小计／总计聚合行，不计入业务明细")

# ---- 数据覆盖披露（动态）：有记录日期区间 + 期间内缺哪几天 ----
_dates_seen = sorted({_iso_day(d[1]) for d in detail if _iso_day(d[1])})
_bad_time = sum(1 for d in detail if not _iso_day(d[1]))
if _bad_time:
    print(f"⚠ 销售流水有 {_bad_time} 行销售时间无法解析，数据覆盖期间将不含这些行，请核对")
if _dates_seen:
    _d0, _d1 = _date.fromisoformat(_dates_seen[0]), _date.fromisoformat(_dates_seen[-1])
    _span = {(_d0 + _timedelta(i)).isoformat() for i in range((_d1 - _d0).days + 1)}
    _gap = sorted(_span - set(_dates_seen))
    _p0 = _date.fromisoformat(str(P_START)[:10])
    _pre = (f"；{_p0.isoformat()} 至 {(_d0 - _timedelta(1)).isoformat()} 无记录" if _d0 > _p0 else "")
    COV_TXT = (f"{_dates_seen[0]} 至 {_dates_seen[-1]}"
               + (f"，其中 " + "、".join(g[5:].replace("-", "/") for g in _gap) + " 无记录" if _gap else "")
               + _pre)
else:
    COV_TXT = "无有效销售记录"

STORE_ORDER = [s for s in ["示例门店一店", "示例门店二店", "示例门店三店", "示例门店四店", "示例门店五店", "示例门店六店"]
               if s in {k for v in oct_store.values() for k in v}]
STORE_ORDER += [k for v in oct_store.values() for k in v if k not in STORE_ORDER]
STORE_ORDER = list(dict.fromkeys(STORE_ORDER))
print(f"识别：库存表「{inv_sheet_name}」/ 次月表「{ws_nov.title if ws_nov is not None else '（无需回加）'}」/ 销售表「{ws_oct.title}」")

# ---------- 合同价格与贴花：唯一取价来源，SCD2 区间错误直接阻断 ----------
def _load_rate_book(wb, sheet_name, value_names, kind, required):
    from settlement_rules import rate_value_column
    try:  # 与 _数据映射.json 读取同款健壮导入：依赖缺失时中文中止，不留半截堆栈
        from schema_common import locate_column, ColumnAmbiguous, ColumnMissing
    except Exception as _col_import_error:  # noqa: BLE001 - 依赖缺失时中文中止
        _bail(f"主档价格／贴花列定位需要 schema_common.py（{_col_import_error}）")

    actual = _master_sheet(sheet_name)
    if actual not in wb.sheetnames:
        if required:
            _bail(f"主档缺少「{sheet_name}」表，无法从合同政策取价")
        return None
    ws = wb[actual]
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    def col(names, required_col=True):
        try:
            return locate_column(headers, names, sheet_name=sheet_name, required=required_col)
        except ColumnMissing as _e:
            _bail(f"主档「{sheet_name}」表缺少列：{'／'.join(str(a) for a in _e.aliases)}"
                  f"；实际表头：{_e.headers}")
        except ColumnAmbiguous as _e:
            _bail(f"主档「{sheet_name}」表「{_e.alias}」命中多列："
                  + "、".join(f"第{i + 1}列「{headers[i]}」" for i in _e.hits)
                  + "；不得任选一列，请只保留一个权威列")
    ci_bar = col(("商品条码", "条码"))
    ci_bar = col(("商品条码", "条码"))
    ci_value = rate_value_column(headers, value_names, actual)
    ci_start = col(("生效起",))
    ci_end = col(("生效止",), False)
    ci_source = col(("依据文件",), False)
    rate_rows = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        if not values or ci_bar >= len(values) or values[ci_bar] in (None, ""):
            continue
        rate_rows.append({
            "barcode": values[ci_bar], "value": values[ci_value], "start": values[ci_start],
            "end": values[ci_end] if ci_end is not None else None,
            "source": values[ci_source] if ci_source is not None else "",
        })
    if required and not rate_rows:
        _bail(f"主档「{sheet_name}」表没有任何有效记录")
    return RateBook(rate_rows, kind=kind)

# 同一主档只完整加载一次：先取价格/贴花两本 RateBook，再取 SKU 表，读完统一关闭。
# 异常归属不变：加载/取价失败仍报价格错，SKU 段失败仍报 SKU 错；finally 保证关闭。
_wb = None
try:
    try:
        _wb = openpyxl.load_workbook(CFG["master_path"], data_only=True)
        PRICE_BOOK = _load_rate_book(_wb, "价格", ("供货价", "结算价", "合同价"), "合同价", True)
        DECAL_BOOK = _load_rate_book(_wb, "贴花", ("贴花标准",), "贴花标准", False)
    except RuleViolation as _e:
        _bail(f"价格政策阻断[{_e.code}] {_e.message}")
    except Exception as _e:
        _bail(f"主档价格／贴花读取失败：{_e}")

    # ---------- 读主档 SKU：别名反查 + 结算属性 + 条码全集 ----------
    MS_BAR, MS_ALIAS, MS_SETTLE, MS_NAME = {}, {}, set(), {}
    try:
        _sku_sheet = _master_sheet("SKU")
        if _sku_sheet in _wb.sheetnames:
            _ws = _wb[_sku_sheet]
            _h = [c.value for c in _ws[1]]
            def _cc(k, d=None):
                # 列定位统一走共享实现：精确优先、多列命中阻断（旧版子串先到先得已废除）。
                try:
                    hit = _locate_column(_h, (k,), sheet_name="主档·SKU", required=False)
                except _ColumnAmbiguous as exc:
                    _bail(f"主档「SKU」表「{k}」命中多列，无法唯一取数：{exc}")
                return d if hit is None else hit
            ci_bar = _cc("条码", 0); ci_name = _cc("标准品名", 1)
            ci_alias = _cc("别名", 2); ci_attr = _cc("结算属性", 6)
            ci_spec = _cc("规格", 3); ci_unit = _cc("单位", 4)
            for r in _ws.iter_rows(min_row=2, values_only=True):
                if not r or not r[ci_bar]:
                    continue
                b = str(r[ci_bar]).strip()
                attr = str(r[ci_attr] or "").strip() if ci_attr is not None else ""
                MS_BAR[b] = dict(name=r[ci_name], attr=attr,
                                 alias=r[ci_alias] if ci_alias is not None else "",
                                 spec=(r[ci_spec] if ci_spec is not None and len(r) > ci_spec else ""),
                                 unit=(r[ci_unit] if ci_unit is not None and len(r) > ci_unit else ""))
                if attr == "结算类":
                    MS_SETTLE.add(b)
                MS_NAME[norm(r[ci_name] or "")] = b
                if ci_alias is not None and r[ci_alias]:
                    for a in str(r[ci_alias]).split("|"):
                        if a.strip():
                            MS_ALIAS[norm(a)] = b
            print(f"主档：SKU {len(MS_BAR)} 条，结算类 {len(MS_SETTLE)} 条，别名 {len(MS_ALIAS)} 条")
    except Exception as _e:
        _bail(f"主档 SKU 表读取失败（{_e}）：结算属性三态校验与条码全集依赖该表，已阻断")
finally:
    if _wb is not None:
        _wb.close()


# ---------- 条码 ↔ 物料编码对照（可选；挂账表只有物料编码时用它兜住匹配）----------
BARCODE_MAP, MAP_SRC = {}, ""
_P_MAP = _src_file("barcode_map", (CFG.get("sources") or {}).get("barcode_map") or "")
if _P_MAP and os.path.exists(_P_MAP):
    _mmx = _read_matrix(_P_MAP, sheet=_src_entry("barcode_map").get("sheet"), role="barcode_map")
    _mhr, _mhd = _src_header_rows(_mmx, "barcode_map", ["条码", "物料编码"])
    m_bar = _src_col("barcode_map", "bar", _mhd, ["商品条码", "条码", "国际条码"], 0,
                     label="对照表·条码列")
    m_code = _src_col("barcode_map", "code", _mhd, ["物料编码", "客户物料号", "物料编号", "商品编码"], 1,
                      label="对照表·物料编码列")
    _dup = []
    for _ri, _r in enumerate(_mmx[_mhr + 1:], _mhr + 2):
        _vb = _cell_of(_r, m_bar)
        _vc = _cell_of(_r, m_code)
        _vb = _norm_bar(_vb)
        _vc = "" if _vc is None else str(_vc).strip()
        if not _vb and not _vc:
            continue
        if not _vb or not _vc:
            print(f"⚠ 对照表第 {_ri} 行条码或物料编码为空，已跳过")
            continue
        if _vc in BARCODE_MAP and BARCODE_MAP[_vc] != _vb:
            _dup.append((_vc, BARCODE_MAP[_vc], _vb))
        else:
            BARCODE_MAP[_vc] = _vb
    if _dup:
        _bail("条码对照表存在一对多：" + "；".join("%s→%s / %s" % x for x in _dup[:5])
              + "。同一物料编码只能对应一个商品条码，请先修正对照表")
    MAP_SRC = os.path.basename(_P_MAP)
    print(f"条码对照：{len(BARCODE_MAP)} 条（{MAP_SRC}）")


def _bar_of(code):
    """物料编码 → 商品条码（对照表未提供或未命中时返回空串）。"""
    return BARCODE_MAP.get(str(code).strip(), "") if code else ""


def resolve_sku(name, code=""):
    """条码优先 → 物料编码对照 → 主档别名反查 → 品名归一化。返回 (库存行, 依据)。"""
    k = norm(name)
    _mb = _bar_of(code)
    if _mb and _mb in inv_by_bar:
        return inv_by_bar[_mb], "物料编码"
    for cand in (oct_by_name.get(k), MS_ALIAS.get(k), MS_NAME.get(k)):
        if cand and cand in inv_by_bar:
            return inv_by_bar[cand], "条码"
    iv = inv_by_n.get(k)
    if iv:
        return iv, "品名"
    return None, "未匹配"

def _shelf_months(value):
    """保质期解析：'36' / '36个月' / '3年' → 月数；识别不了返回 None。"""
    if value in (None, ""):
        return None
    txt = str(value)
    m = re.search(r"(\d+(?:\.\d+)?)", txt)
    if not m:
        return None
    n = float(m.group(1))
    if "年" in txt:
        n *= 12
    elif "天" in txt:
        n = n / 30.0
    return n


# ---------- 批次效期：优先挂账表的批次/有效期列，其次可选批次表 ----------
_P_BATCH = _src_file("batch", (CFG.get("sources") or {}).get("batch") or "")
BATCH_SRC = "甲方挂账表（批次号／有效期列）" if _BATCH_FROM_JIA else ""
BATCH_NOTE = ""
if not BATCH and _P_BATCH and os.path.exists(_P_BATCH):
    _bmx = _read_matrix(_P_BATCH, sheet=_src_entry("batch").get("sheet"), role="batch")
    _bhr, _bhd = _src_header_rows(_bmx, "batch", ["商品名称", "数量"])
    q_bar = _src_col("batch", "bar", _bhd, ["条码", "编码"], 0)
    q_name = _src_col("batch", "name", _bhd, ["商品名称", "名称", "品名"], 1)
    q_lot = _src_col("batch", "lot", _bhd, ["批次号", "批次", "批号"], None)
    q_exp = _src_col("batch", "exp", _bhd, ["有效期", "到期日", "失效"], None)
    q_mfg = _src_col("batch", "mfg", _bhd, ["生产日期", "生产"], None)
    q_days = _src_col("batch", "shelf", _bhd, ["保质期"], None)
    q_qty = _src_col("batch", "qty", _bhd, ["库存数量", "数量", "库存"], None)
    q_code = _src_col("batch", "code", _bhd, ["物料编码", "客户物料号", "物料编号"], None)
    _q_bad = 0
    for _ri, _r in enumerate(_bmx[_bhr + 1:], _bhr + 2):
        _nm = _cell_of(_r, q_name)
        if not _nm:
            continue
        _exp = _parse_day(_cell_of(_r, q_exp)) if q_exp is not None else None
        if _exp is None and q_mfg is not None and q_days is not None:
            _mfg = _parse_day(_cell_of(_r, q_mfg))
            _mon = _shelf_months(_cell_of(_r, q_days))
            if _mfg and _mon:
                _exp = _add_months(_mfg, _mon)
        if _exp is None:
            _q_bad += 1
            continue
        _bb = _norm_bar(_cell_of(_r, q_bar))
        if not _bb and q_code is not None:
            _bb = _bar_of(_cell_of(_r, q_code))
        _lot_v, _lot_exp = _split_lot(_cell_of(_r, q_lot) if q_lot is not None else None)
        if _exp is None:
            _exp = _lot_exp
        if _exp is None:
            _q_bad += 1
            continue
        BATCH.append(dict(name=str(_nm).strip(), bar=_bb,
                          lot=_lot_v or "（未标批次）",
                          exp=_exp,
                          qty=_qval(_cell_of(_r, q_qty), f"批次表第 {_ri} 行·数量"),
                          init=0, price=0.0))
    BATCH_SRC = "甲方批次效期表：" + os.path.basename(_P_BATCH)
    if _q_bad:
        print(f"⚠ 批次表有 {_q_bad} 行取不到有效期（也无法由生产日期＋保质期推算），已跳过，请核对源表")
if not BATCH:
    BATCH_NOTE = ("本期未提供批次效期数据：甲方挂账表无「批次号／有效期」列，"
                  "且未提供批次效期表，故未生成效期预警")

# 批次与促销物料的条码补齐（挂账表行按品名，主档/库存/销售三处反查）
for _b in BATCH:
    if not _b["bar"]:
        _b["bar"] = (_bar_of(_b.get("code", "")) or oct_by_name.get(norm(_b["name"]))
                     or MS_ALIAS.get(norm(_b["name"])) or MS_NAME.get(norm(_b["name"])) or "")
for _k, _p in PROMO.items():
    _p["bar"] = (oct_by_name.get(_k) or MS_ALIAS.get(_k) or MS_NAME.get(_k) or "")

# ---------- 乙方试用装/中小样/配赠使用数据：使用表优先，其次销售流水标记行 ----------
_P_USE = _src_file("promo_use", (CFG.get("sources") or {}).get("promo_use") or "")
PROMO_USE, PROMO_USE_SRC = {}, ""
if _P_USE and os.path.exists(_P_USE):
    _umx = _read_matrix(_P_USE, sheet=_src_entry("promo_use").get("sheet"), role="promo_use")
    _uhr, _uhd = _src_header_rows(_umx, "promo_use", ["商品名称", "数量"])
    u_bar = _src_col("promo_use", "bar", _uhd, ["条码", "编码"], 0)
    u_name = _src_col("promo_use", "name", _uhd, ["商品名称", "名称", "品名"], 1)
    u_qty = _src_col("promo_use", "qty", _uhd, ["使用数量", "领用数量", "数量"], None)
    for _ri, _r in enumerate(_umx[_uhr + 1:], _uhr + 2):
        if not _cell_of(_r, u_name):
            continue
        _ub = _norm_bar(_cell_of(_r, u_bar))
        _uq = _qval(_cell_of(_r, u_qty), f"使用表第 {_ri} 行·数量")
        _uk = _ub or norm(_cell_of(_r, u_name))
        PROMO_USE[_uk] = PROMO_USE.get(_uk, 0) + _uq
    PROMO_USE_SRC = "乙方使用表：" + os.path.basename(_P_USE)
elif _sales_use:
    PROMO_USE = dict(_sales_use)
    PROMO_USE_SRC = "乙方销售流水中的赠品／试用标记行"



# ==================== 核心计算 ====================
rows = []
_missing = []
_uncls = []
if not MS_BAR:
    _bail("主档 SKU 表未读到任何条码，无法判定结算属性与条码全集，已阻断。请核对主档路径与「SKU」表。")
if MS_BAR:
    for x in jia:
        _b = (_bar_of(x.get("code", "")) or oct_by_name.get(norm(x["name"]))
              or MS_ALIAS.get(norm(x["name"])) or MS_NAME.get(norm(x["name"])) or "")
        _a = MS_BAR.get(_b, {}).get("attr") if _b else None
        if _a not in ("结算类", "非结算类"):
            _uncls.append((x["name"], _b or "（无条码）", _a or "空"))
    for x in inv:
        if x["bar"] not in MS_BAR:
            _uncls.append((x["name"], x["bar"] or "（无条码）", "主档未建档"))
        elif MS_BAR[x["bar"]]["attr"] not in ("结算类", "非结算类"):
            _uncls.append((x["name"], x["bar"], MS_BAR[x["bar"]]["attr"] or "空"))
    # 销售独有条码也必须在过滤前检查；未建档不等于非结算，绝不静默漏算。
    for line in sales_rule_lines:
        barcode = line["barcode"]
        attribute = MS_BAR.get(barcode, {}).get("attr")
        if attribute not in ("结算类", "非结算类"):
            reason = "主档未建档" if barcode not in MS_BAR else (attribute or "空")
            _uncls.append((oct_name.get(barcode) or "销售流水商品", barcode or "（无条码）", reason))
if _uncls:
    for _n, _b, _a in _uncls[:12]:
        sys.stderr.write(f"    - {_n}｜{_b}｜结算属性：{_a}\n")
    _bail(f"有 {len(_uncls)} 个商品在主档中未明确「结算类/非结算类」。"
          f"空值不会被默认成结算类，请先补主档 SKU 表再对账")

opening_by_bar, inventory_by_bar = {}, {}
_source_settled_q0 = 0
for x in jia:
    k = norm(x["name"])
    _source_bar = (_bar_of(x.get("code", "")) or oct_by_name.get(k)
                   or MS_ALIAS.get(k) or MS_NAME.get(k) or "")
    # 原始 jia／inv／detail 仍用于来源展示；这里构造的字典与 rows 仅供结算规则使用。
    if MS_BAR.get(_source_bar, {}).get("attr") == "非结算类":
        continue
    iv, _how = resolve_sku(x["name"], x.get("code", ""))
    if iv is not None and _how in ("条码", "物料编码") and norm(iv["name"]) != k:
        print(f"⚠ 挂账「{x['name']}」与库存「{iv['name']}」品名不同，已按条码匹配")
    if iv is None:
        _missing.append(x["name"])
        _b = oct_by_name.get(k) or MS_ALIAS.get(k) or MS_NAME.get(k) or ""
        iv = dict(name=x["name"], bar=_b, spec="", cat="", qty=0, retail=0.0, _absent=True)
    bar = iv["bar"]
    if not bar:
        _bail(f"挂账商品「{x['name']}」无法唯一定位条码，请先补主档或条码映射")
    if bar in opening_by_bar:
        _bail(f"挂账表中同一条码出现多组结算行：{bar}；请先核对，避免动销双算")
    opening_by_bar[bar] = x["qty"]
    _source_settled_q0 += x["qty"]
    if not iv.get("_absent"):
        inventory_by_bar[bar] = iv["qty"]
    _note = ""
    if not iv.get("_absent") and norm(iv["name"]) != k:
        _note = f"乙方系统品名：{iv['name']}"
    if iv.get("_absent"):
        _note = "乙方本期未提供该 SKU 库存，不计算盈亏"
    _ms = MS_BAR.get(bar, {})
    rows.append(dict(name=x["name"], bar=bar,
                     spec=(_ms.get("spec") or iv["spec"]),
                     unit=(_ms.get("unit") or oct_unit.get(bar) or "—"),
                     price=0.0, retail=iv["retail"], q0=x["qty"], q10=0, q11=nov.get(bar, 0),
                     theo=0, act=iv["qty"], adj=0, diff=0, diff_amt=0.0, settle=0.0,
                     init=x["init"],
                     absent=bool(iv.get("_absent")), note=_note))

_rule_sales = [line for line in sales_rule_lines
               if MS_BAR.get(line["barcode"], {}).get("attr") == "结算类"]
_rule_ids = [str(line["id"]) for line in _rule_sales]
if len(_rule_ids) != len(set(_rule_ids)):
    _bail("销售／退货流水号必须全局唯一，才能追溯原销售价和贴花；请修正重复流水号")
_unknown_rule_bars = sorted({line["barcode"] for line in _rule_sales} - set(opening_by_bar))
if _unknown_rule_bars:
    _bail("以下结算条码有销售／退货流水，但不在本次甲方期初导出中："
          + "、".join(_unknown_rule_bars)
          + "。请在甲方导出中补对应行（期初为 0 也要明确列出），不猜测期初数量")
try:
    for _bar in sorted(set(opening_by_bar) | {line["barcode"] for line in _rule_sales}):
        PRICE_BOOK.validate_period(_bar, P_START, P_END)
    RULE_RESULT = calculate_consignment(
        period_start=P_START, period_end=P_END,
        opening=opening_by_bar, inventory=inventory_by_bar,
        addback={bar: nov.get(bar, 0) for bar in opening_by_bar},
        sales=_rule_sales, prices=PRICE_BOOK, decals=DECAL_BOOK,
        original_sales=CFG.get("original_sales") or [],
        variance_authorization=CFG.get("variance_authorization"),
    )
except RuleViolation as _e:
    _bail(f"结算规则阻断[{_e.code}] {_e.message}")

_rule_sku = {item["barcode"]: item for item in RULE_RESULT["quantity_reconciliation"]}
_amount_by_bar = defaultdict(float)
for _line in RULE_RESULT["priced_lines"]:
    _amount_by_bar[_line["barcode"]] += float(_line["amount"])
_priced_by_id = {str(line["id"]): line for line in RULE_RESULT["priced_lines"]}
for _d in detail:
    _priced = _priced_by_id.get(str(_d[0]))
    if _priced and MS_BAR.get(str(_d[4]), {}).get("attr") == "结算类":
        _d[11] = float(_priced["unit_price"])
        _d[12] = float(_priced["amount"])
        _d[13] = _priced["price_evidence"]

def _qty(value):
    """数量统一精度：规则层返回 Decimal，保留实数（如 9.5），不得 int 截断。"""
    if value is None:
        return None
    return round(float(value), 4)


def _qfmt(value):
    """数量显示：整数不带小数点（3），实数保留必要位（9.5），不出现 3.0。"""
    if value is None:
        return "0"
    number = float(value)
    if number == int(number):
        return str(int(number))
    return ("%f" % number).rstrip("0").rstrip(".")


def _qfmt_s(value):
    """带符号数量显示：+3／-3／+9.5。"""
    number = float(value or 0)
    sign = "+" if number > 0 else ""
    return sign + _qfmt(number)


for r in rows:
    _rule = _rule_sku[r["bar"]]
    r["price"] = float(PRICE_BOOK.rate_on(r["bar"], P_END).value)
    r["q10"] = _qty(_rule["net_sales_qty"])
    r["theo"] = _qty(_rule["theoretical_end_qty"])
    r["adj"] = _qty(_rule["adjusted_actual_qty"]) if _rule["adjusted_actual_qty"] is not None else 0
    r["diff"] = _qty(_rule["variance_qty"]) if _rule["variance_qty"] is not None else 0
    r["diff_amt"] = float(_rule["variance_display_amount"] or 0)
    r["settle"] = float(money(_amount_by_bar.get(r["bar"], 0.0)))

T = dict(q0=0, q10=0, q11=0, theo=0, act=0, adj=0, diff=0, diff_amt=0, settle=0)
for r in rows:
    for kk in T:
        if r.get("absent") and kk in ("act", "adj", "diff", "diff_amt", "q11"):
            continue
        T[kk] += r[kk]
T = {k: (float(money(v)) if k in ("diff_amt", "settle") else round(v, 2)) for k, v in T.items()}

OPEN_VAL = float(money(sum(r["q0"] * r["price"] for r in rows)))  # 期末合同价参考估值，非甲方金额
GROSS_SALES = float(RULE_RESULT["gross_sales_amount"])
RETURN_AMOUNT = float(RULE_RESULT["return_amount"])
VARIANCE_INCLUDED = bool(RULE_RESULT["variance_included"])
SETTLEMENT_TOTAL = float(RULE_RESULT["settlement_amount"])

_PART.update(q0=T.get("q0"), open=OPEN_VAL, q10=T.get("q10"),
             settle=T.get("settle"), diff=T.get("diff"), diff_amt=T.get("diff_amt"))
DIFF_RATE = T["diff"] / abs(T["theo"]) if T["theo"] else 0

_prov = [r for r in rows if not r.get("absent")]
THEO_Q_P = sum(r["theo"] for r in _prov)
Q0_P = sum(r["q0"] for r in _prov)
Q10_P = sum(r["q10"] for r in _prov)
for _batch in BATCH:
    _matched = next((r for r in rows if r["bar"] == _batch.get("bar")), None)
    if _matched:
        _batch["price"] = _matched["price"]

# ---------- 效期：期末批次结存（本期动销按有效期升序 FIFO 扣减）----------
_EXP_BASE = _date.fromisoformat(str(P_END)[:10])
_q10_by_bar = {r["bar"]: r["q10"] for r in rows if r["bar"]}
_theo_by_bar = {r["bar"]: r["theo"] for r in rows if r["bar"]}

FIFO_ISSUES = []
_by_bar_lot = defaultdict(list)
for _b in BATCH:
    _by_bar_lot[_b["bar"] or norm(_b["name"])].append(_b)
for _key, _lst in _by_bar_lot.items():
    _lst.sort(key=lambda z: (z["exp"], str(z["lot"])))
    _need = max(_q10_by_bar.get(_key, 0), 0)
    for _b in _lst:
        _take = min(_b["qty"], _need)
        _b["used"] = _take
        _b["left"] = _b["qty"] - _take
        _need -= _take
    if _need > 0:
        print(f"⚠ 「{_lst[0]['name']}」批次数量不足以扣减本期动销，缺口 {_need} 件，请核对批次数据")
    # 系统核销口径：后批次已扣减、而更早批次仍有余额 → 与先进先出不符
    _consumed = [i for i, b in enumerate(_lst) if b["init"] and b["qty"] < b["init"]]
    if _consumed:
        _viol = [b for b in _lst[:min(_consumed)] if b["qty"] > 0]
        if _viol:
            FIFO_ISSUES.append((_lst[0]["name"], [b["lot"] for b in _viol],
                                [b["lot"] for b in _lst[min(_consumed):]
                                 if b["init"] and b["qty"] < b["init"]]))
if FIFO_ISSUES:
    for _n, _earlier, _later in FIFO_ISSUES[:8]:
        print(f"⚠ FIFO 校验：「{_n}」较早批次（{'、'.join(_earlier[:3])}）仍有余额，"
              f"但较晚批次（{'、'.join(_later[:3])}）已发生扣减，存在后进先出迹象")

for _b in BATCH:
    _b["days"] = (_b["exp"] - _EXP_BASE).days
    _b["tier"] = ("已过期" if _b["days"] < 0 else
                  "30 天内" if _b["days"] <= 30 else
                  "90 天内" if _b["days"] <= 90 else
                  "180 天内" if _b["days"] <= 180 else "正常（>180 天）")
EXP_TIERS = ["已过期", "30 天内", "90 天内", "180 天内", "正常（>180 天）"]
EXP_ALERT_TIERS = ("已过期", "30 天内", "90 天内")
EXP_SUM = {t: dict(lots=0, skus=set(), qty=0) for t in EXP_TIERS}
for _b in BATCH:
    if _b["left"] <= 0:
        continue
    _sm = EXP_SUM[_b["tier"]]
    _sm["lots"] += 1
    _sm["skus"].add(_b["bar"] or norm(_b["name"]))
    _sm["qty"] += _b["left"]
EXP_ALERT_QTY = sum(EXP_SUM[t]["qty"] for t in EXP_ALERT_TIERS)
EXP_ALERT_SKU = len({_b["bar"] or norm(_b["name"]) for _b in BATCH
                     if _b["left"] > 0 and _b["tier"] in EXP_ALERT_TIERS})
EXP_TOTAL_LEFT = sum(_b["left"] for _b in BATCH if _b["left"] > 0)

# 勾稽：各条码「期末批次结存」应等于该 SKU 的期末理论结存（q0 − 动销）
EXP_RECON = []
for _key, _lst in _by_bar_lot.items():
    _left = sum(b["left"] for b in _lst)
    _theo = _theo_by_bar.get(_key)
    if _theo is None:
        EXP_RECON.append((_key, _left, None, "甲方挂账表无该条码"))
    elif _left != _theo:
        EXP_RECON.append((_key, _left, _theo, "批次结存≠账面期末结存"))
_exp_diff_qty = sum(abs(_l - (_t or 0)) for _, _l, _t, _ in EXP_RECON)

if _missing:
    print(f"⚠ 以下 {len(_missing)} 个 SKU 乙方本期未提供库存数据，不参与盈亏计算（非按 0 计）：")
    for _m in _missing:
        print(f"    - {_m}")


def _check_identity(label, a, b):
    """数量与各项金额独立复算；不再用调价价差倒挤货值。"""
    if abs(a - b) > 0.01:
        _bail(f"{label}：{a:,.2f} ≠ {b:,.2f}")


if abs(T["q0"] - (T["q10"] + T["theo"])) > 0.01:
    _bail(f"数量勾稽不成立：甲方期初 {_qfmt(T['q0'])} ≠ 本期净动销 {_qfmt(T['q10'])} ＋ 理论结存 {_qfmt(T['theo'])}")
if abs(T["adj"] - (T["act"] + T["q11"])) > 0.01:
    _bail(f"数量勾稽不成立：实际结存 {T['adj']} ≠ 库存 {T['act']} ＋ 回加 {T['q11']}")
if abs(T["diff"] - (THEO_Q_P - T["adj"])) > 0.01:
    _bail(f"数量勾稽不成立：盈亏 {_qfmt(T['diff'])} ≠ 可核对理论结存 {_qfmt(THEO_Q_P)} － 实际结存 {_qfmt(T['adj'])}")
_check_identity("逐笔销售／退货货款复算", T["settle"], float(RULE_RESULT["sales_settlement_amount"]))
_check_identity("盈亏列示金额复算", T["diff_amt"], float(RULE_RESULT["variance_display_amount"]))
_expected_total = T["settle"] + (T["diff_amt"] if VARIANCE_INCLUDED else 0.0)
_check_identity("本期结算总额复算", SETTLEMENT_TOTAL, _expected_total)

diff_rows   = [r for r in rows if not r.get("absent") and abs(r["diff"]) > 1e-9]
ok_rows     = [r for r in rows if not r.get("absent") and abs(r["diff"]) <= 1e-9 and r["q0"] > 0]
zero_rows   = [r for r in rows if not r.get("absent") and abs(r["q0"]) <= 1e-9]
absent_rows = [r for r in rows if r.get("absent")]
_jia_bars_all = {r["bar"] for r in rows if r["bar"]}
_jia_names_all = {norm(r["name"]) for r in rows}
extra     = [x for x in inv
             if not ((x["bar"] and x["bar"] in _jia_bars_all) or norm(x["name"]) in _jia_names_all)]
EXTRA_Q   = sum(x["qty"] for x in extra)

# ==================== 贴花：正常销售与退货冲减逐笔计算 ====================
TH_ROWS, TH_ACTIVE = [], DECAL_BOOK is not None
TH_TOTAL = float(RULE_RESULT["decal_amount"])
_th_by_bar = {}
for _line in RULE_RESULT["priced_lines"]:
    if DECAL_BOOK is None:
        break
    _bar = _line["barcode"]
    _item = _th_by_bar.setdefault(_bar, {"qty": 0, "amt": 0.0, "standards": set()})
    _item["qty"] += round(float(_line["quantity"]), 4)
    _item["amt"] += float(_line["decal_amount"])
    if _line.get("decal_standard") is not None:
        _item["standards"].add(float(_line["decal_standard"]))
for r in rows:
    _item = _th_by_bar.get(r["bar"])
    if not _item or (not _item["qty"] and abs(_item["amt"]) < 0.005):
        continue
    _standards = _item["standards"]
    TH_ROWS.append(dict(name=r["name"], bar=r["bar"], qty=_item["qty"],
                        std=next(iter(_standards)) if len(_standards) == 1 else None,
                        mixed=len(_standards) > 1, amt=float(money(_item["amt"]))))
_check_identity("贴花逐笔独立复算", TH_TOTAL, sum(x["amt"] for x in TH_ROWS))
if TH_ACTIVE:
    print(f"贴花：{len(TH_ROWS)} 个动销 SKU，合计 {TH_TOTAL:,.2f} 元（正常销售取 {P_END} 标准，退货按原销售标准）")
else:
    print("贴花：主档无「贴花」表，本期不计算销售奖励")
print(f"输入目录：{BASE}")
print(f"源文件：{F_JIA} / {F_OCT} / {F_INV}")
print(f"输出文件：{OUT}")
print(f"甲方期初 {_qfmt(T['q0'])} 件 | 本期净动销 {_qfmt(T['q10'])} 件／流水货款 {T['settle']:,.2f} 元 | 回加 {_qfmt(T['q11'])} 件")
print(f"理论结存 {_qfmt(T['theo'])} 件 | 库存导出 {_qfmt(T['act'])} 件＋回加＝实际结存 {_qfmt(T['adj'])} 件 | 盈亏 {_qfmt(T['diff'])} 件／{T['diff_amt']:,.2f} 元")
if MS_BAR:
    _inv_bars = {x["bar"] for x in inv if x["bar"]}
    _unreg = _inv_bars - set(MS_BAR)
    if _unreg:
        print(f"⚠ 库存表有 {len(_unreg)} 个条码未在主档 SKU 登记：{sorted(_unreg)[:6]}")
    _jia_bars = {}
    for x in jia:
        _b = (oct_by_name.get(norm(x["name"])) or MS_ALIAS.get(norm(x["name"]))
              or MS_NAME.get(norm(x["name"])) or "")
        if _b:
            _jia_bars[_b] = _jia_bars.get(_b, 0) + 1
    _dupb = [b for b, n in _jia_bars.items() if n > 1]
    if _dupb:
        _bail(f"挂账表中同一条码出现多行（{_dupb[:3]}），请先在主档/挂账表核对，避免动销双算")
    if MS_SETTLE and len(_jia_bars) and len(_jia_bars) != len(MS_SETTLE):
        print(f"⚠ 挂账覆盖条码 {len(_jia_bars)} 个，主档结算类 {len(MS_SETTLE)} 个，差额 "
              f"{sorted(MS_SETTLE - set(_jia_bars))[:5]}，请确认是否漏挂/多挂")

_nounit = [r["name"] for r in rows if r["unit"] == "—"]
if _nounit:
    print(f"⚠ {len(_nounit)} 个单品未在主档登记单位，对账单按「—」显示：{_nounit[:5]}")

if not ADD_BACK:
    # 库存文件里若存在含数据的次月销售表，说明库存实际导出于期末之后——asof 大概率填错
    _late_q = 0
    try:
        _wl = [w for w in wb1.worksheets if w is not ws_inv and w.max_row > 1]
        if _wl:
            _lw = max(_wl, key=lambda w: w.max_row)
            for _r in _lw.iter_rows(min_row=2, values_only=True):
                for _v in _r:
                    if isinstance(_v, (int, float)) and _v:
                        _late_q += float(_v)
    except Exception:
        pass
    if _late_q > 0:
        _bail(f"库存文件中存在含 {_late_q} 件数量记录的「次月销售」类工作表，"
              f"说明库存实际导出于期末 {P_END} 之后；但 --asof={ASOF} 不晚于期末，"
              f"本次按无需回加处理会虚增短缺。请把 --asof 改为真实库存导出日（次月 2 日左右）")

_det_sum = round(sum(float(d[7] or 0) for d in detail), 4)
if abs(_det_sum - sum(oct_qty.values())) > 1e-9:
    _bail(f"明细合计 {_det_sum} ≠ 按条码汇总 {sum(oct_qty.values())}（源表可能有合计行或空行）")
_src_q0 = _source_settled_q0   # 促销物料与主档非结算类均不参与结算勾稽
if abs(_src_q0 - T["q0"]) > 0.01:
    _bail(f"挂账源表数量合计 {_src_q0} ≠ 归一化后 {T['q0']}（同品名多行合并异常）")

_unreg_rows = [x for x in extra if (x["bar"] not in MS_BAR)] if MS_BAR else list(extra)
if _unreg_rows:
    _q = sum(x["qty"] for x in _unreg_rows)
    print(f"⚠ 库存有 {len(_unreg_rows)} 行、{_q} 件在主档 SKU 中查不到条码，已归入非结算桶，请建档或确认：")
    for x in _unreg_rows[:10]:
        print(f"    - {x['bar'] or '（无条码）'} {x['name']} {_qfmt(x['qty'])} 件")
elif MS_BAR and EXTRA_Q:
    print(f"非结算桶 {EXTRA_Q} 件，均已在主档标为非结算类")

_sum_q0 = sum(x["q0"] for x in rows)
if _sum_q0 != T["q0"]:
    _bail(f"甲方期初数量合计 {_sum_q0} ≠ {T['q0']}")
print(f"盈亏 {_qfmt(T['diff'])} 件／{T['diff_amt']:,.2f} 元 | 数量勾稽与金额分列复算通过 | 非结算物料 {EXTRA_Q} 件")

# ==================== 工作簿框架 ====================
out = openpyxl.Workbook(); out.remove(out.active)

def setup(ws, widths, orient="portrait"):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.sheet_view.showGridLines = False
    ws.print_options.gridLines = False
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = orient
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.35
    ws.page_margins.top = ws.page_margins.bottom = 0.45

def sec(ws, r, c1, c2, text):
    for c in range(c1, c2 + 1):
        ws.cell(row=r, column=c).fill = fill(LBLUE)
        ws.cell(row=r, column=c).border = Border(bottom=MED)
    cell = ws.cell(row=r, column=c1, value=text)
    cell.font = F(10.5, True, NAVY)
    cell.alignment = A("left", "center", False, 1)
    ws.row_dimensions[r].height = 26

def hdr(ws, r, c1, headers, h=36, fl=NAVY):
    for i, t in enumerate(headers):
        cell = ws.cell(row=r, column=c1 + i, value=t)
        cell.font = F(8.5, True, WHITE)
        cell.fill = fill(fl)
        cell.alignment = A("center", "center", True)
        cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=MED)
    ws.row_dimensions[r].height = h

def num(cell, v, fmt="#,##0"):
    if isinstance(v, float) and v == int(v):
        v = int(v)  # 整数值不写 60.0，数量列保持整数观感；实数仍保留小数
    cell.value = v; cell.number_format = fmt
    cell.font = F(9)
    cell.alignment = A("right", "center", False, 1) if ".00" in fmt else A("center")

# ==================== 00 对账说明 ====================
ws = out.create_sheet("00-对账说明")
setup(ws, [10.5] * 12)
ws.row_dimensions[1].height = 8
mput(ws, 2, 1, 2, 12, f"{P_CN}  铺货代销对账报告", F(22, True, WHITE), A("center"), NAVY)
ws.row_dimensions[2].height = 48
mput(ws, 3, 1, 3, 12, f"{PA_NAME}  ×  {PB_NAME}      |      对账期间 {P_START} 至 {P_END}",
     F(9.5, False, NAVY), A("center"), LBLUE)
ws.row_dimensions[3].height = 24
ws.row_dimensions[4].height = 6

sec(ws, 5, 1, 12, "一、对账基础信息")
info = [("甲方（供货方）", PA_NAME, "乙方（代销方）", PB_NAME),
        ("甲方纳税人识别号", PA_TAX, "乙方纳税人识别号", PB_TAX),
        ("经销合同编号", CONTRACT_NO, "结算方式", SETTLE_TERMS_TXT),
        ("对账期间", f"{P_START} 至 {P_END}", "数据基准日", ASOF)]
r = 6
for lab1, v1, lab2, v2 in info:
    mput(ws, r, 1, r, 2, lab1, F(9, True, DGREY), A("left", "center", True, 1), GREY)
    mput(ws, r, 3, r, 6, v1, F(9.5), A("left", "center", True, 1))
    mput(ws, r, 7, r, 8, lab2, F(9, True, DGREY), A("left", "center", True, 1), GREY)
    mput(ws, r, 9, r, 12, v2, F(9.5), A("left", "center", True, 1))
    ws.row_dimensions[r].height = 24; r += 1
mput(ws, r, 1, r, 2, "数据来源", F(9, True, DGREY), A("left", "center", True, 1), GREY)
mput(ws, r, 3, r, 12, f"甲方未结算数据（{ASOF} 导出）／乙方销售数据（{_d_end.year} 年 {_d_end.month} 月）／乙方库存数据（{ASOF}）",
     F(9.5), A("left", "center", True, 1))
ws.row_dimensions[r].height = 24; r += 2

sec(ws, r, 1, 12, "二、对账结论"); r += 1
_variance_label = "盈亏列示（净短缺，元）" if T["diff"] > 0 else ("盈亏列示（净盘盈，元）" if T["diff"] < 0 else "盈亏列示（元）")
kpi = [("本期结算总额（元）", SETTLEMENT_TOTAL, f"净动销 {_qfmt(T['q10'])} 件 · {'含本期授权盈亏' if VARIANCE_INCLUDED else '盈亏未纳入'}", NAVY, VLBLUE),
       ("正常销售货款（元）", GROSS_SALES, "逐笔销售数量 × 销售日合同价", BLUE, VLBLUE),
       ("退货冲减金额（元）", RETURN_AMOUNT, "按原销售价冲减，不按退货日猜价", GREEN, LGREEN),
       (_variance_label, T['diff_amt'],
        (f"{_qfmt_s(T['diff'])} 件 · 正数短缺／负数盘盈" if diff_rows else f"{len(ok_rows)} 个单品账实相符"), RED, LRED)]
for i, (lab, val, sub, col, bg) in enumerate(kpi):
    c1 = 1 + i * 3
    mput(ws, r, c1, r, c1 + 2, lab, F(8.5, True, col), A("center"), bg)
    cc = mput(ws, r + 1, c1, r + 1, c1 + 2, val, F(16.5, True, col), A("center"), bg)
    cc.number_format = '"¥ "#,##0.00'
    mput(ws, r + 2, c1, r + 2, c1 + 2, sub, F(7.5, False, DGREY), A("center"), bg)
ws.row_dimensions[r].height = 21
ws.row_dimensions[r + 1].height = 38
ws.row_dimensions[r + 2].height = 22
r += 4

sec(ws, r, 1, 12, "三、结算与核对口径（数量勾稽、金额分列）"); r += 1
formula = [
    ("① 正常销售货款", "＝ Σ（每笔销售数量 × 该笔销售日期生效合同价）"),
    ("② 退货冲减金额", "＝ －Σ（每笔退货数量 × 原销售单价）；数量在本期同步冲减"),
    ("③ 理论期末数量", f"＝ 本次甲方期初数量 － 本期净动销数量【截至 {P_END}】"),
    ("④ 盘点盈亏", "＝ 理论期末 －（库存导出＋导出日前已销回加）；正数短缺、负数盘盈"),
    ("⑤ 金额分列", "盈亏按期末合同价列示；默认不进总额，纳入须绑定本期的单次授权，不用价差倒挤"),
]
for lab, txt in formula:
    mput(ws, r, 1, r, 3, lab, F(9, True, BLUE), A("left", "center", True, 1), VLBLUE)
    mput(ws, r, 4, r, 12, txt, F(9), A("left", "center", True, 1))
    ws.row_dimensions[r].height = 24; r += 1
_add_note = (
    f"说明：甲方期初数量 {T['q0']} 件直接采用本次甲方导出，不与上期台账强制衔接。乙方库存导出日为 {ASOF}，"
    f"晚于对账截止日 {P_END}，其间已发生销售 {T['q11']} 件；回加后统一至 {P_END} 核对。" if ADD_BACK else
    f"说明：甲方期初数量 {T['q0']} 件直接采用本次甲方导出，不与上期台账强制衔接。乙方库存导出日 {ASOF} 不晚于期末，无需回加。")
mput(ws, r, 1, r, 12, _add_note, F(8.5, False, AMBER), A("left", "center", True, 1), LAMBER)
ws.row_dimensions[r].height = 46; r += 2

sec(ws, r, 1, 12, "四、主要结论"); r += 1
concl = [
    f"1. 正常销售货款 {GROSS_SALES:,.2f} 元，退货冲减 {RETURN_AMOUNT:,.2f} 元，流水结算金额 {T['settle']:,.2f} 元；本期结算总额 {SETTLEMENT_TOTAL:,.2f} 元。",
    (f"2. 库存核对：{len(ok_rows) + len(diff_rows)} 个有铺货单品中 {len(ok_rows)} 个账实相符；{len(diff_rows)} 个存在盈亏，净盈亏 {_qfmt_s(T['diff'])} 件／{T['diff_amt']:,.2f} 元（正数短缺、负数盘盈）。"
     if diff_rows else f"2. 库存核对：{len(ok_rows)} 个有铺货单品全部账实相符，盈亏为 0。"),
    ("3. 逐 SKU 盈亏：" + "、".join(f"{r['name']} {_qfmt_s(r['diff'])} 件（{'短缺' if r['diff'] > 0 else '盘盈'}）" for r in sorted(diff_rows, key=lambda x: -x['diff']))
     + "；原因与归属待双方核对，本报告不作责任认定。") if diff_rows else
    "3. 逐 SKU 库存与账面盈亏均为 0，明细见附表 02《库存差异表》。",
    (f"4. 有 {len(absent_rows)} 个 SKU 未提供乙方库存，不按 0 计，已单独列为待补依据。" if absent_rows else
     "4. 全部结算 SKU 均已提供乙方库存，任何非零盈亏均已逐 SKU 列示，不设比例阈值。"),
    (f"5. 盈亏列示金额 {T['diff_amt']:,.2f} 元"
     + ("已依据仅绑定本期的单次授权纳入结算总额。" if VARIANCE_INCLUDED else "未纳入本期结算总额；如需纳入，须另行对本期间单次授权。")) if diff_rows else
    "5. 本期库存账实相符，无盈亏需核对。",
    f"6. 合同价格政策是货款单价唯一来源；甲方仅提供期初数量。正常销售按销售日取价，退货按原销售价冲减。",
] + ([f"7. 贴花费用合计 {TH_TOTAL:,.2f} 元：正常销售按 {P_END} 生效标准，退货按原销售标准冲减；盘点盈亏不参与贴花，兑付方式由双方书面确认。"] if TH_ACTIVE else [])
if V17:
    _open_all = sum(_ledger_sum(k, "opening") for k in _ledger_kinds())
    concl.append(f"{len(concl) + 1}. 试用装／中小样：期初均为 {_open_all:,} 件，"
                 f"本期期间发货{_kinds_phrase('shipped')}；"
                 f"乙方使用{_kinds_phrase('used')}；"
                 f"期末结存{_kinds_phrase('closing')}。"
                 f"{_policy_clause()}，明细见附表 08。")
    if EXPIRY_ON:
        concl.append(f"{len(concl) + 1}. 商品批次效期提示：在库商品按剩余效期分档列示，"
                     f"明细见附表 09。")
else:
    if BATCH:
        concl.append(f"{len(concl) + 1}. 批次效期：期末 {len([_b for _b in BATCH if _b['left'] > 0])} 个批次有结存，其中 {EXP_ALERT_SKU} 个 SKU、{EXP_ALERT_QTY:,} 件"
                     f"将在 90 天内到期或已过期（含已过期），明细见附表 08。"
                     + ("" if not FIFO_ISSUES else f"另有 {len(FIFO_ISSUES)} 个 SKU 存在后进先出迹象，见附表 08。"))
    else:
        concl.append(f"{len(concl) + 1}. 批次效期：本期未提供批次效期数据，未生成效期预警；"
                     f"请在挂账表补「批次号／有效期至」列后重跑（见附表 08）。")
    if PROMO:
        concl.append(f"{len(concl) + 1}. 试用装／中小样／配赠：本期甲方出库 {sum(_p['out_qty'] for _p in PROMO.values()):,} 件，"
                     + ("乙方已提供使用数据，差异见附表 09。"
                        if PROMO_USE_SRC else "乙方未提供使用数据，按出库口径列示（未经乙方确认）。")
                     + "该部分不参与货款结算。")
for t in concl:
    mput(ws, r, 1, r, 12, t, F(9), A("left", "center", True, 1))
    ws.row_dimensions[r].height = _fit_h(t)
    r += 1
r += 1

sec(ws, r, 1, 12, "五、需乙方书面确认／说明的事项"); r += 1
hdr(ws, r, 1, ["序号", "事  项", "", "", "涉 及 内 容", "", "", "", "需 乙 方 回 复", "", "", ""], 26)
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=8)
ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=12)
r += 1
_asks_raw = [("本期结算总额确认", f"{SETTLEMENT_TOTAL:,.2f} 元／净动销 {_qfmt(T['q10'])} 件，逐笔核对销售日合同价与原退货价", f"签章确认，确认后 {CT_SIGN_TXT}内寄回原件")]
if diff_rows:
    _asks_raw.append(("库存盈亏核对", f"净 {_qfmt_s(T['diff'])} 件／列示 {T['diff_amt']:,.2f} 元（正数短缺、负数盘盈），逐 SKU 核对并提供支持单据", f"收到后 {CT_OBJ_TXT}内书面回复；是否纳入结算须逐期单次授权"))
_asks_raw += [("非结算类货品说明", f"{len(extra)} 个单品共 {EXTRA_Q:,} 件，未列入甲方挂账台账，不属甲方铺货货品，不纳入本期结算范围", "上述货品归属另行确定，不参与本期及后续货款结算；请乙方核对数量"),
              ("结算期间划分", (f"{NEXT_LABEL}销售 {T['q11']} 件仅回加库存、不计入本期"
                               if ADD_BACK else "库存导出日不晚于期末，无导出日前销售需回加"), "请书面回复确认")]
if V17:
    _asks_raw.append(("试用装／中小样使用数量",
                      f"期初 {_open_all:,} 件；"
                      f"期间发货 {_kinds_phrase('shipped')}；"
                      f"使用 {_kinds_phrase('used')}；"
                      f"期末 {_kinds_phrase('closing')}",
                      "请签章确认上述使用数量与期末数量；如有差异请附明细说明"))
    if EXPIRY_ON:
        _asks_raw.append(("近效期库存处理", "在库商品按剩余效期分档列示（附表 09）",
                          "请确认近效期商品的动销安排；处理方式由双方另行书面协商"))
else:
    if BATCH and EXP_ALERT_QTY:
        _asks_raw.append(("近效期库存处理",
                          f"{EXP_ALERT_SKU} 个 SKU、{EXP_ALERT_QTY:,} 件将在 90 天内到期或已过期（附表 08）",
                          "请优先动销或书面申请退回，处理方式与费用承担由双方确认"))
    elif not BATCH:
        _asks_raw.append(("批次效期数据", "本期未提供批次／有效期数据，无法预警",
                          "请在挂账表补「批次号／有效期至」列，或提供批次效期表"))
    if PROMO and not PROMO_USE_SRC:
        _asks_raw.append(("试用装／配赠使用数据",
                          f"甲方出库 {sum(_p['out_qty'] for _p in PROMO.values()):,} 件，乙方未提供使用数量",
                          "请补报使用明细，或书面确认按出库计列"))
    if FIFO_ISSUES:
        _asks_raw.append(("批次核销口径",
                          f"{len(FIFO_ISSUES)} 个 SKU 存在后进先出迹象（较早批次仍有余额、较晚批次已扣减）",
                          "请确认批次核销是否按先进先出执行"))
asks = [(str(_i), _a, _b, _c) for _i, (_a, _b, _c) in enumerate(_asks_raw, 1)]
for no, item, cont, rep in asks:
    mput(ws, r, 1, r, 1, no, F(9, True), A("center"))
    mput(ws, r, 2, r, 4, item, F(9, True), A("left", "center", True, 1))
    mput(ws, r, 5, r, 8, cont, F(9), A("left", "center", True, 1))
    mput(ws, r, 9, r, 12, rep, F(9), A("left", "center", True, 1))
    ws.row_dimensions[r].height = 32; r += 1
r += 1

sec(ws, r, 1, 12, "六、附表清单"); r += 1
hdr(ws, r, 1, ["附  表", "", "用  途", "", "", "", "", "", "", "", "", ""], 24)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=12)
r += 1
_annex = [("01", "对账单", "正式结算单据，可直接打印盖章，作为乙方确认及付款依据"),
          ("02", "库存差异表", ("逐 SKU 账实核对（含导出日前销量回加），作为乙方核对及书面说明依据" if ADD_BACK
                              else "逐 SKU 账实核对，作为乙方核对及书面说明依据")),
          ("03", "源数据对照", "甲方未结算数据／乙方动销／乙方库存 三源逐 SKU 汇总对照，供双方逐行确认"),
          ("04", "甲方挂账明细", "甲方未结算数据（截至基准日），期初余额依据"),
          ("05", "乙方销售汇总", f"{_d_end.month} 月动销按 SKU × 门店交叉汇总，销量核对依据"),
          ("06", "乙方库存明细", f"乙方 {_d_asof:%Y%m%d} 库存数据，库存核对依据"),
          ("07", "乙方销售流水", f"乙方当期销售流水明细 {len(detail)} 行，逐笔溯源")]
if V17:
    _annex.append(("08", SHEET_MAP.get("08", "试用装与中小样"),
                   "试用装／中小样期初、期间发货、使用与期末数量，供乙方核对确认"))
    if EXPIRY_ON:
        _annex.append(("09", SHEET_MAP.get("09", "效期预警"),
                       "在库商品批次、到期日与剩余效期，按 32 / 28 / 24 个月分档提示"))
else:
    _annex += [("08", "效期预警", ("按批次列示剩余天数与分档，供乙方优先动销" if BATCH
                                  else "本期未提供批次效期数据，仅出说明页")),
               ("09", "试用装与配赠", ("试用装／中小样／配赠出库与使用情况，供费用核算"
                                     if PROMO else "本期挂账表无促销物料行"))]
for no, nm, use in _annex:
    mput(ws, r, 1, r, 2, f"附表 {no}", F(9, True, BLUE), A("center"), VLBLUE)
    mput(ws, r, 3, r, 4, nm, F(9, True), A("left", "center", True, 1))
    mput(ws, r, 5, r, 12, use, F(9), A("left", "center", True, 1))
    ws.row_dimensions[r].height = 21; r += 1
r += 1

sec(ws, r, 1, 12, "七、签章确认"); r += 1
mput(ws, r, 1, r, 6, f"甲方（盖章）：{PA_NAME}", F(9.5, True), A("left", "center", True, 1))
mput(ws, r, 7, r, 12, f"乙方（盖章）：{PB_NAME}", F(9.5, True), A("left", "center", True, 1))
ws.row_dimensions[r].height = 30; r += 1
mput(ws, r, 1, r, 6, "授权代表（签字）：", F(9), A("left", "center", True, 1))
mput(ws, r, 7, r, 12, "授权代表（签字）：", F(9), A("left", "center", True, 1))
ws.row_dimensions[r].height = 42; r += 1
mput(ws, r, 1, r, 6, "日　期：　　　年　　月　　日", F(9), A("left", "center", True, 1))
mput(ws, r, 7, r, 12, "日　期：　　　年　　月　　日", F(9), A("left", "center", True, 1))
ws.row_dimensions[r].height = 30
print("00 ok")

# ==================== 01 对账单 ====================
ws = out.create_sheet("01-对账单")
setup(ws, [6, 42, 12, 7, 12, 14, 16])
ws.row_dimensions[1].height = 10
mput(ws, 2, 1, 2, 7, "销  售  对  账  单", F(22, True, NAVY), A("center"))
ws.row_dimensions[2].height = 44
mput(ws, 3, 1, 3, 7, f"{P_CN}    ·    对账期间：{P_START} 至 {P_END}", F(10, False, DGREY), A("center"))
ws.row_dimensions[3].height = 22
for c in range(1, 8):
    ws.cell(row=4, column=c).border = Border(top=MED)
ws.row_dimensions[4].height = 6

mput(ws, 5, 1, 5, 4, f"甲方（供货方）：{PA_NAME}", F(9.5, True), A("left", "center", True, 0))
mput(ws, 5, 5, 5, 7, f"单据编号：{STATEMENT_NO or '＿＿＿＿＿＿'}　　制单日期：{ASOF_CN}", F(9, False, DGREY), A("right", "center", True, 0))
mput(ws, 6, 1, 6, 4, f"纳税人识别号：{PA_TAX}", F(9, False, DGREY), A("left", "center", True, 0))
mput(ws, 6, 5, 6, 7, f"合同编号：{CONTRACT_NO}", F(9, False, DGREY), A("right", "center", True, 0))
mput(ws, 7, 1, 7, 4, f"乙方（代销方）：{PB_NAME}", F(9.5, True), A("left", "center", True, 0))
mput(ws, 7, 5, 7, 7, "币种：人民币（CNY）", F(9, False, DGREY), A("right", "center", True, 0))
mput(ws, 8, 1, 8, 4, f"纳税人识别号：{PB_TAX}", F(9, False, DGREY), A("left", "center", True, 0))
mput(ws, 8, 5, 8, 7, f"结算方式：{CT_METHOD_TXT}", F(9, False, DGREY), A("right", "center", True, 0))
for rr in range(5, 9):
    for c in range(1, 8):
        ws.cell(row=rr, column=c).border = Border(bottom=Side(style="dotted", color=BORDC))
    ws.row_dimensions[rr].height = 21
mput(ws, 9, 1, 9, 7, f"结算依据：按合同约定（{SETTLE_TERMS_TXT}）计算；已发货但截至对账日尚未动销的铺货产品，不计入本期结算货款（合同约定的结算口径）。",
     F(8.5, False, BLUE), A("left", "center", True, 0), VLBLUE)
ws.row_dimensions[9].height = 26
ws.row_dimensions[10].height = 5

sec(ws, 11, 1, 7, "一、本期结算明细")
hdr(ws, 12, 1, ["序号", "商品名称", "规格", "单位", "期末参考单价\n（元）", "本期净动销数量", "流水结算金额\n（元）"], 34)
r = 13
idx = 0
pay_rows = [x for x in rows if x["q0"] > 0 or abs(x["q10"]) > 1e-9 or abs(x["settle"]) > 0.005]
for x in pay_rows:
    idx += 1
    bg = WHITE if idx % 2 else GREY
    mput(ws, r, 1, r, 1, idx, F(9), A("center"), bg)
    mput(ws, r, 2, r, 2, x["name"], F(9), A("left", "center", True, 0), bg)
    mput(ws, r, 3, r, 3, x["spec"], F(9), A("center"), bg)
    mput(ws, r, 4, r, 4, x["unit"], F(9), A("center"), bg)
    c = ws.cell(row=r, column=5); num(c, x["price"], "#,##0.00"); c.fill = fill(bg)
    c = ws.cell(row=r, column=6); num(c, x["q10"]); c.fill = fill(bg)
    if abs(x["q10"]) <= 1e-9:
        c.font = F(9, c="A6A6A6")
    c = ws.cell(row=r, column=7); num(c, x["settle"], "#,##0.00"); c.fill = fill(bg)
    ws.row_dimensions[r].height = 26
    r += 1
mput(ws, r, 1, r, 4, "本期结算总额", F(9.5, True, NAVY), A("center"), LBLUE)
c = ws.cell(row=r, column=5); c.fill = fill(LBLUE); c.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
c = ws.cell(row=r, column=6); num(c, T["q10"]); c.font = F(10, True, NAVY)
c.fill = fill(LBLUE); c.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
c = ws.cell(row=r, column=7); num(c, SETTLEMENT_TOTAL, "#,##0.00"); c.font = F(10, True, NAVY)
c.fill = fill(LBLUE); c.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
ws.row_dimensions[r].height = 28
r += 1
mput(ws, r, 1, r, 3, "人民币（大写）", F(9, True, DGREY), A("center"), GREY)
mput(ws, r, 4, r, 7, rmb_upper(SETTLEMENT_TOTAL), F(10.5, True, NAVY), A("left", "center", True, 1), GREY)
ws.row_dimensions[r].height = 30
r += 1
mput(ws, r, 1, r, 7, "说明：期末参考单价只用于库存盈亏列示，不代表本 SKU 的逐笔成交均价；每笔试算价与金额见附表 07。",
     F(8, False, DGREY), A("left", "center", True, 1), GREY)
r += 2

sec(ws, r, 1, 7, "二、金额分列（数量勾稽见附表 02）"); r += 1
hdr(ws, r, 1, ["项目", "", "金额（元）", "计算／说明"], 26)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
r += 1
roll = [
    ("① 正常销售货款", GROSS_SALES, "Σ（每笔销售数量 × 销售日合同价）", False),
    ("② 退货冲减金额", RETURN_AMOUNT, "－Σ（退货数量 × 原销售单价）", False),
    ("③ 流水结算金额", T["settle"], "＝ ① ＋ ②；等于第一节各 SKU 流水金额之和", True),
    ("④ 盘点盈亏列示金额", T["diff_amt"], f"{_qfmt_s(T['diff'])} 件 × 各 SKU 期末合同价；正数短缺、负数盘盈", False),
    ("⑤ 本期结算总额", SETTLEMENT_TOTAL,
     (("＝ ③ ＋ ④（已获仅绑定本期的单次授权）；与第一节合计及大写金额一致" if VARIANCE_INCLUDED
       else "＝ ③；④ 仅列示核对，未纳入总额；与第一节合计及大写金额一致")), True),
]
for lab, amt, note, bold in roll:
    mput(ws, r, 1, r, 2, lab, F(9, bool(bold), NAVY if bold else "000000"), A("left", "center", True, 1), LBLUE if bold else WHITE)
    c = ws.cell(row=r, column=3); num(c, amt, "#,##0.00")
    c.font = F(9.5, bool(bold), "000000")
    c.fill = fill(LBLUE if bold else WHITE)
    mput(ws, r, 4, r, 7, note, F(8.5, False, DGREY), A("left", "center", True, 1), LBLUE if bold else WHITE)
    ws.row_dimensions[r].height = 24
    r += 1
r += 1

sec(ws, r, 1, 7, "三、贴花费用（销售奖励，独立于货款与盈亏）"); r += 1
if TH_ACTIVE:
    hdr(ws, r, 1, ["序号", "商品名称", "净动销数量", "贴花标准\n（元/件）", "贴花金额\n（元）", "依据", ""], 24)
    ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
    r += 1
    for i, x in enumerate(TH_ROWS, 1):
        mput(ws, r, 1, r, 1, i, F(9), A("center"), WHITE)
        mput(ws, r, 2, r, 2, x["name"], F(9), A("left", "center", True, 1), WHITE)
        c = ws.cell(row=r, column=3); num(c, x["qty"]); c.font = F(9); c.fill = fill(WHITE)
        c = ws.cell(row=r, column=4)
        if x["std"] is None:
            c.value = "逐笔"; c.font = F(8.5); c.alignment = A("center"); c.fill = fill(WHITE)
        else:
            num(c, x["std"], "#,##0.00"); c.font = F(9); c.fill = fill(WHITE)
        c = ws.cell(row=r, column=5); num(c, x["amt"], "#,##0.00"); c.font = F(9); c.fill = fill(WHITE)
        mput(ws, r, 6, r, 7, "正常销售取期末标准；退货按原销售标准冲减", F(8, False, DGREY), A("left", "center", True, 1), WHITE)
        ws.row_dimensions[r].height = 20
        r += 1
    mput(ws, r, 1, r, 2, "合　计", F(9.5, True, WHITE), A("center"), NAVY)
    c = ws.cell(row=r, column=3); num(c, sum(x["qty"] for x in TH_ROWS)); c.font = F(9.5, True, WHITE)
    c.fill = fill(NAVY); c.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    c = ws.cell(row=r, column=4); c.value = None; c.fill = fill(NAVY); c.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    c = ws.cell(row=r, column=5); num(c, TH_TOTAL, "#,##0.00"); c.font = F(10, True, WHITE)
    c.fill = fill(NAVY); c.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    mput(ws, r, 6, r, 7, "正常销售与退货冲减逐笔相加", F(8, False, WHITE), A("left", "center", True, 1), NAVY)
    ws.row_dimensions[r].height = 26
    r += 1
    mput(ws, r, 2, r, 7,
         f"口径：正常销售按数量 × {P_END} 生效贴花标准；原销售已计提的退货按原标准冲减，未计提不冲。"
         f"盘点盈亏不参与贴花；兑付方式与开票口径按合同约定执行。",
         F(8.5, False, AMBER), A("left", "center", True, 1), LAMBER)
    ws.row_dimensions[r].height = 44
    r += 1
else:
    mput(ws, r, 2, r, 7, "本期未计算销售奖励。",
         F(8.5, False, DGREY), A("left", "center", True, 1), WHITE)
    r += 1
r += 1
sec(ws, r, 1, 7, "四、对账说明"); r += 1
notes = [
    ("1.", f"本对账单依据《经销合同》（编号 {CONTRACT_NO}）约定的结算口径出具（{SETTLE_TERMS_TXT}），未动销铺货不计入本期。"),
    ("2.", f"数据来源：本次甲方数量导出（{ASOF}）、合同价格政策、乙方 {_d_end.year} 年 {_d_end.month} 月销售数据及 {ASOF} 库存数据。甲方单价与金额不参与结算。"),
    ("3.", (f"乙方库存导出日为 {ASOF}，{NEXT_LABEL}期间发生销售 {T['q11']} 件，仅回加库存核对，不计入本期结算；归属期间由双方确认。"
             if ADD_BACK else f"乙方库存导出日为 {ASOF}，不晚于对账截止日 {P_END}，无需回加导出日前销售。")),
    ("4.", (f"数量勾稽：{len(diff_rows)} 个 SKU 存在盈亏，净盈亏 {_qfmt_s(T['diff'])} 件／列示 {T['diff_amt']:,.2f} 元（正数短缺、负数盘盈）。"
            + ("本期间已获单次授权，盈亏纳入结算总额。" if VARIANCE_INCLUDED else "盈亏未纳入本期结算总额，任何非零项均待双方核对。")
            if diff_rows else f"数量勾稽：{len(ok_rows)} 个有铺货单品全部账实相符，无盈亏（详见附表 02）。")),
    ("5.", f"乙方应于收到本对账单后 {CT_OBJ_TXT}内书面提出异议，逾期未提出视为确认；确认后 {CT_SIGN_TXT}内将签章原件寄回甲方。"),
    ("6.", f"付款期限：按合同原文及主档登记的付款截止日执行；本报告不自行因节假日提前或顺延。部分异议不影响无争议款项如期支付。"),
    ("7.", "合同价格政策为单价唯一权威源；期中调价逐笔按业务日期取价，退货按原销售价冲减；开票事宜按合同约定执行。"),
]
if not V17:
    if PROMO:
        notes.append((f"{len(notes) + 1}.",
                      f"试用装／中小样／配赠本期出库 {sum(_p['out_qty'] for _p in PROMO.values()):,} 件，"
                      f"不参与货款结算；使用情况见附表 09（乙方"
                      f"{'已' if PROMO_USE_SRC else '未'}提供使用数据）。"))
    if BATCH and EXP_ALERT_QTY:
        notes.append((f"{len(notes) + 1}.",
                      f"批次效期：{EXP_ALERT_SKU} 个 SKU、{EXP_ALERT_QTY:,} 件将在 90 天内到期或已过期，"
                      f"请乙方优先动销或申请退回；明细见附表 08。"))
for no, t in notes:
    mput(ws, r, 1, r, 1, no, F(8.5, True, BLUE), A("right", "top", False, 0))
    mput(ws, r, 2, r, 7, t, F(8.5), A("left", "center", True, 0))
    ws.row_dimensions[r].height = _fit_h(t)
    r += 1
for t in [f"户　名：{BK_NAME}",
          f"账　号：{BK_NO}",
          f"开户行：{BK_BANK}"]:
    mput(ws, r, 2, r, 7, t, F(8.5, True, "000000"), A("left", "center", True, 1), VLBLUE)
    ws.row_dimensions[r].height = 18; r += 1
r += 2

sec(ws, r, 1, 7, "五、双方确认"); r += 1
mput(ws, r, 1, r, 3, "甲方（盖章）：", F(9, True), A("left", "center", True, 1))
mput(ws, r, 4, r, 7, "乙方（盖章）：", F(9, True), A("left", "center", True, 1))
ws.row_dimensions[r].height = 46; r += 1
mput(ws, r, 1, r, 3, "授权代表（签字）：", F(9), A("left", "center", True, 1))
mput(ws, r, 4, r, 7, "授权代表（签字）：", F(9), A("left", "center", True, 1))
ws.row_dimensions[r].height = 34; r += 1
mput(ws, r, 1, r, 3, "日　期：　　　年　　月　　日", F(9), A("left", "center", True, 1))
mput(ws, r, 4, r, 7, "日　期：　　　年　　月　　日", F(9), A("left", "center", True, 1))
ws.row_dimensions[r].height = 26
ws.print_area = f"A1:G{r}"
print("01 ok")

# ==================== 02 库存差异表 ====================
ws = out.create_sheet("02-库存差异表")
setup(ws, [4, 13, 28, 8, 4, 9, 10, 9, 9, 9, 9, 10, 9, 10, 8, 17], "landscape")
ws.row_dimensions[1].height = 10
mput(ws, 2, 1, 2, 16, "库 存 差 异 表", F(20, True, NAVY), A("center"))
ws.row_dimensions[2].height = 38
mput(ws, 3, 1, 3, 16, f"{PA_NAME}  ×  {PB_NAME}    ·    对账期间 {P_START} 至 {P_END}    ·    库存数据基准日 {ASOF}",
     F(9, False, DGREY), A("center"))
ws.row_dimensions[3].height = 22
mput(ws, 4, 1, 4, 16,
     (f"口径：乙方库存导出日为 {ASOF}，晚于对账截止日 {P_END}，须将 {NEXT_LABEL}已发生销售 {T['q11']} 件回加至乙方库存。"
      f"实际结存＝库存导出＋回加；盈亏＝理论期末－实际结存，正数短缺、负数盘盈。任何非零盈亏均列示，不设阈值。"
      if ADD_BACK else
      f"口径：乙方库存导出日 {ASOF} 不晚于期末，无需回加。盈亏＝理论期末－库存导出，正数短缺、负数盘盈；任何非零盈亏均列示。"),
     F(8.5, False, AMBER), A("left", "center", True, 0), LAMBER)
ws.row_dimensions[4].height = 34
ws.row_dimensions[5].height = 6

HDR2 = ["序号", "商品条码", "商品名称", "规格", "单位", "期末合同价\n（列示用）",
        "甲方期初\n数量", f"减：{_d_end.month}月\n净动销", "理论期末\n数量",
        f"乙方 {_d_asof.month:02d}-{_d_asof.day:02d}\n库存（系统导出）",
        (f"加：{_d_next.month}月\n{_d_next.day}—{_d_asof.day}销售" if ADD_BACK else "导出日前销售\n（无需回加）"),
        "实际结存\n数量",
        "盈亏数量\n（＋短缺）", "盈亏列示金额\n（元）", "盈亏率\n（÷理论期末）", "核对状态"]

r = 6
_SEC = ["一", "二", "三", "四", "五"]
_SI = [0]
def _ns():
    _v = _SEC[_SI[0]]; _SI[0] += 1; return _v
if not diff_rows:
    sec(ws, r, 1, 16, f"{_ns()}、库存差异明细"); r += 1
    mput(ws, r, 1, r, 16, f"本期 {len(ok_rows)} 个有铺货单品全部账实相符，无库存差异。",
         F(9, True, GREEN), A("center", "center", True, 1), LGREEN)
    ws.row_dimensions[r].height = 28
    r += 2
else:
    sec(ws, r, 1, 16, f"{_ns()}、库存差异明细（{len(diff_rows)} 个 SKU，需乙方书面说明）"); r += 1
    hdr(ws, r, 1, HDR2, 40); r += 1
    tot_d = dict(q0=0, q10=0, theo=0, act=0, q11=0, adj=0, diff=0, amt=0)
    for i, x in enumerate(diff_rows, 1):
        bg = LRED
        _rate = x["diff"] / x["theo"] if x["theo"] else 0.0
        vals = [i, x["bar"], x["name"], x["spec"], x["unit"], x["price"],
                x["q0"], -x["q10"] if x["q10"] else 0, x["theo"],
                x["act"], x["q11"], x["adj"], x["diff"], x["diff_amt"], _rate]
        for j, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=j)
            if j in (1, 3, 4, 5):
                cell.value = v; cell.font = F(8.5, j == 3); cell.alignment = A("left" if j == 3 else "center", "center", True, 1 if j == 3 else 0)
            elif j == 2:
                cell.value = v; cell.font = F(8.5); cell.alignment = A("center")
                cell.number_format = "@"
            elif j == 15:
                cell.value = abs(v); cell.number_format = "0.00%"; cell.font = F(8.5, True, RED); cell.alignment = A("center")
            else:
                num(cell, v, "#,##0.00" if j in (6, 14) else "#,##0")
                if j in (13, 14): cell.font = F(8.5, True, RED)
            cell.fill = fill(bg)
        mput(ws, r, 16, r, 16, "库存短缺，需说明去向" if x["diff"] > 0 else "库存盘盈，需说明来源", F(8, True, RED), A("center"), bg)
        ws.row_dimensions[r].height = 30
        for k, v in [("q0", x["q0"]), ("q10", x["q10"]), ("theo", x["theo"]), ("act", x["act"]),
                     ("q11", x["q11"]), ("adj", x["adj"]), ("diff", x["diff"]), ("amt", x["diff_amt"])]:
            tot_d[k] += v
        r += 1
    mput(ws, r, 1, r, 5, "小　计", F(9, True, NAVY), A("center"), LBLUE)
    for j, v in [(6, None), (7, tot_d["q0"]), (8, -tot_d["q10"]), (9, tot_d["theo"]), (10, tot_d["act"]),
                 (11, tot_d["q11"]), (12, tot_d["adj"]), (13, tot_d["diff"]), (14, tot_d["amt"])]:
        cell = ws.cell(row=r, column=j)
        if v is None:
            cell.value = None
        else:
            num(cell, v, "#,##0.00" if j == 14 else "#,##0")
            cell.font = F(9, True, NAVY)
        cell.fill = fill(LBLUE)
        cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    tot_rate = tot_d["diff"] / tot_d["theo"] if tot_d["theo"] else 0
    cell = ws.cell(row=r, column=15); cell.value = abs(tot_rate); cell.number_format = "0.00%"
    cell.font = F(9, True, RED); cell.alignment = A("center"); cell.fill = fill(LBLUE)
    cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    mput(ws, r, 16, r, 16, ((f"{_qfmt(abs(tot_d['diff']))} 件净短缺" if tot_d["diff"] > 0 else f"{_qfmt(abs(tot_d['diff']))} 件净盘盈")
                               if tot_d["diff"] else "无差异"), F(9, True, RED), A("center"), LBLUE)
    ws.row_dimensions[r].height = 26
    r += 2

sec(ws, r, 1, 16, f"{_ns()}、核对相符明细（{len(ok_rows)} 个 SKU，账实一致）"); r += 1
hdr(ws, r, 1, HDR2, 40, fl=BLUE); r += 1
tot_o = dict(q0=0, q10=0, theo=0, act=0, q11=0, adj=0)
for i, x in enumerate(ok_rows, 1):
    bg = WHITE if i % 2 else GREY
    vals = [i, x["bar"], x["name"], x["spec"], x["unit"], x["price"],
            x["q0"], -x["q10"] if x["q10"] else 0, x["theo"], x["act"], x["q11"], x["adj"],
            0, 0.0, 0.0]
    for j, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=j)
        if j in (1, 3, 4, 5):
            cell.value = v; cell.font = F(8.5, j == 3); cell.alignment = A("left" if j == 3 else "center", "center", True, 1 if j == 3 else 0)
        elif j == 2:
            cell.value = v; cell.font = F(8.5); cell.alignment = A("center"); cell.number_format = "@"
        elif j in (13, 14, 15):
            num(cell, v if j != 15 else v, "#,##0" if j == 13 else ("#,##0.00" if j == 14 else "0.0%"))
            cell.font = F(8.5, c=GREEN)
        else:
            num(cell, v, "#,##0.00" if j == 6 else "#,##0")
        cell.fill = fill(bg)
    mput(ws, r, 16, r, 16, "✓ 账实相符", F(8, True, GREEN), A("center"), bg)
    ws.row_dimensions[r].height = 26
    for k, v in [("q0", x["q0"]), ("q10", x["q10"]), ("theo", x["theo"]), ("act", x["act"]),
                 ("q11", x["q11"]), ("adj", x["adj"])]:
        tot_o[k] += v
    r += 1
mput(ws, r, 1, r, 5, "小　计", F(9, True, NAVY), A("center"), LBLUE)
for j, v in [(7, tot_o["q0"]), (8, -tot_o["q10"]), (9, tot_o["theo"]), (10, tot_o["act"]),
             (11, tot_o["q11"]), (12, tot_o["adj"]), (13, 0), (14, 0.0), (15, 0.0)]:
    cell = ws.cell(row=r, column=j)
    num(cell, v, "#,##0.00" if j == 14 else "#,##0")
    if j == 15: cell.number_format = "0.0%"
    cell.font = F(9, True, NAVY)
    cell.fill = fill(LBLUE); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
mput(ws, r, 16, r, 16, "全部相符", F(9, True, GREEN), A("center"), LBLUE)
ws.row_dimensions[r].height = 26
r += 1
mput(ws, r, 1, r, 5, "合　计（一＋二）", F(9.5, True, WHITE), A("center"), NAVY)
for j, v in [(6, None), (7, Q0_P), (8, -Q10_P), (9, THEO_Q_P), (10, T["act"]),
             (11, T["q11"]), (12, T["adj"]), (13, T["diff"]), (14, T["diff_amt"]), (15, DIFF_RATE)]:
    cell = ws.cell(row=r, column=j)
    if v is None:
        cell.value = None
    else:
        num(cell, abs(v) if j == 15 else v, "#,##0.00" if j == 14 else "#,##0")
        if j == 15: cell.number_format = "0.00%"
        if j in (13, 14, 15): cell.font = F(9.5, True, RED)
        else: cell.font = F(9.5, True, WHITE)
    cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
mput(ws, r, 16, r, 16, ((f"净短缺 {_qfmt(T['diff'])} 件（{abs(DIFF_RATE)*100:.2f}%）" if T['diff'] > 0
                          else f"净盘盈 {_qfmt(abs(T['diff']))} 件（{abs(DIFF_RATE)*100:.2f}%）") if T['diff'] else "无盈亏"),
     F(9, True, WHITE), A("center"), NAVY)
ws.row_dimensions[r].height = 28
r += 1
_zero_names = [z["name"] for z in zero_rows]
mput(ws, r, 1, r, 16,
     (("注：" + "、".join(_zero_names) + " 本期期初铺货数量为 0，不参与数量差异比对，已列入附表 03、06 核对。")
      if _zero_names else "注：本期无期初铺货为 0 的新品，差异明细覆盖全部有铺货 SKU。"),
     F(8, False, DGREY), A("left", "center", True, 0))
ws.row_dimensions[r].height = 22
r += 2
if absent_rows:
    sec(ws, r, 1, 16, f"{_ns()}、乙方未提供库存的 SKU（{len(absent_rows)} 个，不计入差异合计）"); r += 1
    hdr(ws, r, 1, ["序号", "商品条码", "商品名称", "规格", "单位", "期末合同参考价\n（元）",
                   "甲方期初\n数量", "本期净动销\n（件）", "说　　明"], 32, fl="C55A11")
    ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=16)
    r += 1
    for i, x in enumerate(absent_rows, 1):
        bg = WHITE if i % 2 else GREY
        mput(ws, r, 1, r, 1, i, F(8.5), A("center"), bg)
        _c = ws.cell(row=r, column=2); _c.value = x["bar"]; _c.font = F(8.5)
        _c.alignment = A("center"); _c.number_format = "@"; _c.fill = fill(bg)
        mput(ws, r, 3, r, 3, x["name"], F(8.5), A("left", "center", True, 0), bg)
        mput(ws, r, 4, r, 4, x["spec"], F(8.5), A("center"), bg)
        mput(ws, r, 5, r, 5, x["unit"], F(8.5), A("center"), bg)
        _c = ws.cell(row=r, column=6); num(_c, x["price"], "#,##0.00"); _c.fill = fill(bg)
        _c = ws.cell(row=r, column=7); num(_c, x["q0"]); _c.fill = fill(bg)
        _c = ws.cell(row=r, column=8); num(_c, x["q10"]); _c.fill = fill(bg)
        mput(ws, r, 9, r, 16, "乙方本期未提供该 SKU 库存，差异无法计算；不按 0 计，待乙方补报后重算",
             F(8), A("left", "center", True, 0), bg)
        ws.row_dimensions[r].height = 22
        r += 1
    r += 1
print("02 part ok")

sec(ws, r, 1, 16, f"{_ns()}、非结算类货品（{len(extra)} 个单品，共 {EXTRA_Q:,} 件，未列入甲方挂账台账，不纳入本期结算）"); r += 1
hdr(ws, r, 1, ["序号", "商品条码", "商品名称", "规格", "单位", "零售价\n（元）", "库存数量", "说　　明"], 32, fl="7F7F7F")
ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=16)
r += 1
for i, x in enumerate(extra, 1):
    bg = WHITE if i % 2 else GREY
    mput(ws, r, 1, r, 1, i, F(8.5), A("center"), bg)
    cell = ws.cell(row=r, column=2); cell.value = x["bar"]; cell.font = F(8.5); cell.alignment = A("center"); cell.number_format = "@"; cell.fill = fill(bg)
    mput(ws, r, 3, r, 3, x["name"], F(8.5), A("left", "center", True, 0), bg)
    mput(ws, r, 4, r, 4, x["spec"], F(8.5), A("center"), bg)
    mput(ws, r, 5, r, 5, "片", F(8.5), A("center"), bg)
    cell = ws.cell(row=r, column=6); num(cell, x["retail"], "#,##0.00"); cell.fill = fill(bg)
    cell = ws.cell(row=r, column=7); num(cell, x["qty"]); cell.fill = fill(bg)
    note = "未列入甲方挂账台账，不属甲方铺货货品；归属另行确定，数量以乙方库存为准"
    mput(ws, r, 8, r, 16, note, F(8.5, False, DGREY), A("left", "center", True, 1), bg)
    ws.row_dimensions[r].height = 22
    r += 1
mput(ws, r, 1, r, 5, "小　计", F(9, True, NAVY), A("center"), LBLUE)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
cell = ws.cell(row=r, column=6); cell.value = None; cell.fill = fill(LBLUE)
cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
cell = ws.cell(row=r, column=7); num(cell, EXTRA_Q); cell.font = F(9, True, NAVY); cell.fill = fill(LBLUE)
cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
mput(ws, r, 8, r, 16, "不参与本期及后续结算，仅核对数量", F(8.5, False, DGREY), A("left", "center", True, 1), LBLUE)
ws.row_dimensions[r].height = 24
r += 2

sec(ws, r, 1, 16, f"{_ns()}、差异结论"); r += 1
advice = [
    ("盈亏性质", (f"{len(diff_rows)} 个 SKU 存在盈亏，净盈亏 {_qfmt_s(T['diff'])} 件／列示 {T['diff_amt']:,.2f} 元（正数短缺、负数盘盈）；"
                f"{len(ok_rows)} 个 SKU 账实相符。所有非零项均已列示，无比例阈值。") if diff_rows
               else f"{len(ok_rows)} 个有铺货 SKU 全部账实相符。"),
    ("盈亏分布", ("逐 SKU 盈亏：" + "、".join(f"{r['name']} {_qfmt_s(r['diff'])} 件" for r in sorted(diff_rows, key=lambda x: -x['diff']))
                + f"；其余 {len(ok_rows)} 个单品盈亏为 0。") if diff_rows else
               f"{len(ok_rows)} 个有铺货单品全部账实相符。"),
] + ([
    ("差异原因", "本表仅列示双方数据的逐 SKU 盈亏，不作成因推断或责任认定；原因与归属由甲乙双方核对确认。"),
    ("待确认事项", f"盈亏列示金额 {T['diff_amt']:,.2f} 元"
                 + ("已依据本期单次授权纳入结算总额，授权不得向其他期间继承。" if VARIANCE_INCLUDED
                    else "未纳入本期结算总额；如双方确认需要纳入，须对本期间另行单次授权后重新出单。")),
] if diff_rows else [])
for lab, txt in advice:
    mput(ws, r, 1, r, 2, lab, F(9, True, BLUE), A("center", "center", True, 0), VLBLUE)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    mput(ws, r, 3, r, 16, txt, F(9), A("left", "center", True, 0))
    ws.row_dimensions[r].height = _fit_h(txt)
    r += 1
r += 1

mput(ws, r, 1, r, 8, f"甲方（盖章）：{PA_NAME}", F(9, True), A("left", "center", True, 1))
mput(ws, r, 9, r, 16, f"乙方（盖章）：{PB_NAME}", F(9, True), A("left", "center", True, 1))
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=16)
ws.row_dimensions[r].height = 34; r += 1
mput(ws, r, 1, r, 8, "对账人（签字）：　　　　　　　日期：　　年　　月　　日", F(9), A("left", "center", True, 1))
mput(ws, r, 9, r, 16, "对账人（签字）：　　　　　　　日期：　　年　　月　　日", F(9), A("left", "center", True, 1))
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=16)
ws.row_dimensions[r].height = 30
print("02 ok")

# ==================== 03 源数据对照 ====================
ws = out.create_sheet("03-源数据对照")
setup(ws, [4, 13, 27, 12, 8, 4, 9, 8, 8, 10, 8, 8, 8, 8, 7, 9, 8], "landscape")
ws.row_dimensions[1].height = 10
mput(ws, 2, 1, 2, 17, "源 数 据 对 照 表（甲方挂账 × 乙方动销 × 乙方库存）", F(18, True, NAVY), A("center"))
ws.row_dimensions[2].height = 36
mput(ws, 3, 1, 3, 17, f"供双方逐行确认　·　数量单位：件／盒／瓶／套　·　金额单位：人民币元　·　数据基准日 {ASOF}",
     F(9, False, DGREY), A("center"))
ws.row_dimensions[3].height = 22
ws.row_dimensions[4].height = 6

HDR3 = ["序号", "商品条码", "商品名称（乙方系统）", "甲方系统品名差异", "规格", "单位",
        "期末合同\n参考价", "甲方期初\n数量", f"{_d_end.month} 月\n净动销", "流水结算金额",
        (f"{_d_next.month} 月 {_d_next.day}—{_d_asof.day}\n销售" if ADD_BACK else "导出日前\n销售（无）"),
        f"乙方 {_d_asof.month:02d}-{_d_asof.day:02d}\n库存（系统导出）", "实际结存\n（含回加）", "理论期末\n数量",
        "盈亏数量\n（＋短缺）", "盈亏列示金额", "核对结果"]
r = 5
hdr(ws, r, 1, HDR3, 42); r += 1
for i, x in enumerate(rows, 1):
    _ab = bool(x.get("absent"))
    bg = LAMBER if _ab else (LRED if x["diff"] > 0 else (LAMBER if x["diff"] else (WHITE if i % 2 else GREY)))
    vals = [i, x["bar"], x["name"], x["note"] or "—", x["spec"], x["unit"], x["price"],
            x["q0"], x["q10"], x["settle"], x["q11"], x["act"], x["adj"], x["theo"],
            x["diff"], x["diff_amt"]]
    for j, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=j)
        if j in (1, 5, 6):
            cell.value = v; cell.font = F(8.5); cell.alignment = A("center")
        elif j == 2:
            cell.value = v; cell.font = F(8.5); cell.alignment = A("center"); cell.number_format = "@"
        elif j == 3:
            cell.value = v; cell.font = F(8.5); cell.alignment = A("left", "center", True, 1)
        elif j == 4:
            cell.value = v; cell.font = F(7.5, c=DGREY); cell.alignment = A("left", "center", True, 1)
        elif _ab and j in (12, 13, 15, 16):
            cell.value = "—"; cell.font = F(8.5, c="A6A6A6"); cell.alignment = A("center")
        else:
            num(cell, v, "#,##0.00" if j in (7, 10, 16) else "#,##0")
            cell.font = F(8.5, False, RED if (j in (15, 16) and not _ab) else "000000")
        cell.fill = fill(bg)
    if _ab:
        res, _rc = "未提供库存", "7F7F7F"
    else:
        res = ("✗ 短缺" if x["diff"] > 0 else "✗ 盘盈") if x["diff"] else "✓ 相符"
        _rc = (RED if x["diff"] > 0 else AMBER) if x["diff"] else GREEN
    mput(ws, r, 17, r, 17, res, F(8.5, True, _rc), A("center"), bg)
    ws.row_dimensions[r].height = 24
    r += 1
mput(ws, r, 1, r, 6, "合　计", F(9.5, True, WHITE), A("center"), NAVY)
for j, v in [(7, None), (8, T["q0"]), (9, T["q10"]), (10, T["settle"]), (11, T["q11"]),
             (12, T["act"]), (13, T["adj"]), (14, T["theo"]), (15, T["diff"]), (16, T["diff_amt"])]:
    cell = ws.cell(row=r, column=j)
    if v is None: cell.value = None
    else:
        num(cell, v, "#,##0.00" if j in (10, 16) else "#,##0")
        cell.font = F(9, True, RED if j in (15, 16) else WHITE)
    cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
mput(ws, r, 17, r, 17, f"差异 {abs(DIFF_RATE)*100:.2f}%", F(9, True, WHITE), A("center"), NAVY)
ws.row_dimensions[r].height = 28
r += 2
_has_name_note = any(x["note"] for x in rows)
mput(ws, r, 1, r, 17,
     ("注 1：甲方系统品名与乙方系统品名存在括号全半角、空格或后缀差异，已按商品条码一一对应，品名差异见「甲方系统品名差异」列。"
      if _has_name_note else "注 1：甲方系统品名与乙方系统品名已按商品条码一一对应。"),
     F(8, False, DGREY), A("left", "center", True, 0))
r += 1
mput(ws, r, 1, r, 17,
     (f"注 2：实际结存 ＝ 乙方库存 ＋ {NEXT_LABEL}销售（{T['q11']} 件）；理论期末 ＝ 本次甲方期初 － {_d_end.month} 月净动销。二者同处 {P_END} 时点；盈亏＝理论期末－实际结存，正数短缺、负数盘盈。"
      if ADD_BACK else
      f"注 2：实际结存 ＝ 乙方库存（导出日不晚于期末，无需回加）；理论期末 ＝ 本次甲方期初 － {_d_end.month} 月净动销。二者同处 {P_END} 时点；盈亏＝理论期末－实际结存，正数短缺、负数盘盈。"),
     F(8, False, DGREY), A("left", "center", True, 0))
r += 1
mput(ws, r, 1, r, 17,
     (("注 3：" + "、".join(x["name"] for x in zero_rows) + " 本期期初铺货与库存均为 0，无发生额，已列入本表但不参与结算。")
      if zero_rows else "注 3：本期无期初铺货为 0 的新品，本表覆盖全部有铺货 SKU。"),
     F(8, False, DGREY), A("left", "center", True, 0))
ws.freeze_panes = "D6"
print("03 ok")

# ==================== 04 甲方数量明细 ====================
ws = out.create_sheet("04-甲方挂账明细")
setup(ws, [5, 36, 13, 13, 15, 6, 13, 15, 25, 22], "landscape")
ws.row_dimensions[1].height = 10
mput(ws, 2, 1, 2, 10, "附 表 04　甲方期初数量明细", F(15, True, NAVY), A("center"))
ws.row_dimensions[2].height = 32
mput(ws, 3, 1, 3, 10, f"数量来源：甲方 {ASOF} 导出　·　价格来源：合同价格政策（甲方单价／金额不读取）",
     F(8.5, False, DGREY), A("center"))
ws.row_dimensions[3].height = 20
ws.row_dimensions[4].height = 6
r = 5
hdr(ws, r, 1, ["序号", "商品名称", "甲方期初数量", "初始数量", "商品条码", "单位", "期末参考单价", "期末参考货值", "价格依据", "备注"], 34)
r += 1
for i, x in enumerate(rows, 1):
    bg = WHITE if i % 2 else GREY
    mput(ws, r, 1, r, 1, i, F(9), A("center"), bg)
    mput(ws, r, 2, r, 2, x["name"], F(9), A("left", "center", True, 0), bg)
    for j, v in [(3, x["q0"]), (4, x["init"])]:
        cell = ws.cell(row=r, column=j); num(cell, v); cell.fill = fill(bg)
    cell = ws.cell(row=r, column=5); cell.value = x["bar"]; cell.font = F(9); cell.alignment = A("center"); cell.number_format = "@"; cell.fill = fill(bg)
    mput(ws, r, 6, r, 6, x["unit"], F(9), A("center"), bg)
    cell = ws.cell(row=r, column=7); num(cell, x["price"], "#,##0.00"); cell.fill = fill(bg)
    cell = ws.cell(row=r, column=8); num(cell, float(money(x["q0"] * x["price"])), "#,##0.00"); cell.fill = fill(bg)
    mput(ws, r, 9, r, 9, PRICE_BOOK.rate_on(x["bar"], P_END).source or "合同价格政策", F(8, c=DGREY), A("left", "center", True, 1), bg)
    note = ("本期甲方期初为 0" if abs(x["q0"]) <= 1e-9
            else ("初始数量与甲方期初数量不同，请核对" if abs(x["init"] - x["q0"]) > 1e-9 else "—"))
    mput(ws, r, 10, r, 10, note, F(8, c=DGREY), A("left", "center", True, 1), bg)
    ws.row_dimensions[r].height = 24
    r += 1
mput(ws, r, 1, r, 2, "合　计", F(9.5, True, WHITE), A("center"), NAVY)
for j, v in [(3, T["q0"]), (4, sum(x["init"] for x in rows)), (5, None), (6, None), (7, None), (8, OPEN_VAL), (9, None), (10, None)]:
    cell = ws.cell(row=r, column=j)
    if v is None:
        cell.value = None; cell.font = F(9, True, WHITE)
    else:
        num(cell, v, "#,##0.00" if j == 8 else "#,##0"); cell.font = F(9.5, True, WHITE)
    cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
ws.row_dimensions[r].height = 28
r += 2
_init_same = all(x["init"] == x["q0"] for x in rows)
mput(ws, r, 1, r, 10,
     (f"说明：{len(rows)} 行初始数量与甲方期初数量一致。" if _init_same else "说明：部分初始数量与甲方期初数量不同，已逐行标注。")
     + "本表期初直接采用本次甲方导出，不与上期台账期末强制衔接；参考货值只按期末合同价估值，不参与货款倒挤。",
     F(8, False, DGREY), A("left", "center", True, 0))
ws.row_dimensions[r].height = 34
r += 2
sec(ws, r, 1, 10, "二、数量与价格来源核对"); r += 1
verify = [
    ("数量合计", f"{len(rows)} 行甲方期初数量合计 {T['q0']:,} 件，与数量勾稽表一致。"),
    ("价格来源", "正常销售按销售日合同价逐笔计算；本表期末参考价只用于盈亏和库存参考估值。"),
    ("非结算物料", f"另有 {len(_jia_promo)} 行促销物料，不参与货款结算，见附表 {'08' if V17 else '09'}。"),
]
for lab, txt in verify:
    mput(ws, r, 1, r, 2, lab, F(8.5, True, BLUE), A("center", "center", True, 0), VLBLUE)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    mput(ws, r, 3, r, 10, txt, F(8.5), A("left", "center", True, 0))
    ws.row_dimensions[r].height = 22 if len(txt) < 60 else 34
    r += 1
print("04 ok")

# ==================== 05 乙方销售汇总 ====================
ws = out.create_sheet("05-乙方销售汇总")
setup(ws, [14, 34, 10, 5] + [11] * len(STORE_ORDER) + [8, 7], "landscape")
ws.row_dimensions[1].height = 10
mput(ws, 2, 1, 2, 10, f"附 表 05　乙方 {PERIOD} 动销汇总（按 SKU × 门店）", F(15, True, NAVY), A("center"))
ws.row_dimensions[2].height = 32
mput(ws, 3, 1, 3, 10, f"数据来源：乙方提供的 {_d_end.year} 年 {_d_end.month} 月销售数据　·　共 {len(detail)} 行 / {sum(oct_qty.values())} 件　·　金额单位：人民币元",
     F(8.5, False, DGREY), A("center"))
ws.row_dimensions[3].height = 20
ws.row_dimensions[4].height = 6
r = 5
sec(ws, r, 1, 10, f"一、{P_CN}动销（按门店）"); r += 1
hdr(ws, r, 1, ["商品条码", "商品名称", "规格", "单位"] + STORE_ORDER + ["合　计", "占比"], 32)
_total_col = 5 + len(STORE_ORDER)
_ratio_col = _total_col + 1
r += 1
oct_bars = [b for b in oct_qty]
oct_bars.sort(key=lambda b: -oct_qty[b])
total_oct = sum(oct_qty.values())
for i, b in enumerate(oct_bars, 1):
    bg = WHITE if i % 2 else GREY
    cell = ws.cell(row=r, column=1); cell.value = b; cell.font = F(9); cell.alignment = A("center")
    cell.number_format = "@"; cell.fill = fill(bg)
    mput(ws, r, 2, r, 2, oct_name[b], F(9), A("left", "center", True, 0), bg)
    mput(ws, r, 3, r, 3, oct_spec[b], F(9), A("center"), bg)
    mput(ws, r, 4, r, 4, oct_unit[b], F(9), A("center"), bg)
    for k, st in enumerate(STORE_ORDER):
        cell = ws.cell(row=r, column=5 + k)
        num(cell, oct_store[b].get(st, 0)); cell.fill = fill(bg)
        if oct_store[b].get(st, 0) == 0: cell.font = F(9, c="BFBFBF")
    cell = ws.cell(row=r, column=_total_col); num(cell, oct_qty[b]); cell.font = F(9, True); cell.fill = fill(bg)
    cell = ws.cell(row=r, column=_ratio_col); cell.value = oct_qty[b] / total_oct
    cell.number_format = "0.0%"; cell.font = F(9); cell.alignment = A("center"); cell.fill = fill(bg)
    ws.row_dimensions[r].height = 24
    r += 1
mput(ws, r, 1, r, 4, "合　计", F(9.5, True, WHITE), A("center"), NAVY)
for k, st in enumerate(STORE_ORDER):
    cell = ws.cell(row=r, column=5 + k)
    num(cell, sum(oct_store[b].get(st, 0) for b in oct_bars)); cell.font = F(9.5, True, WHITE)
    cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
cell = ws.cell(row=r, column=_total_col); num(cell, total_oct); cell.font = F(10, True, WHITE)
cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
cell = ws.cell(row=r, column=_ratio_col); cell.value = 1.0; cell.number_format = "0.0%"
cell.font = F(9.5, True, WHITE); cell.alignment = A("center")
cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
ws.row_dimensions[r].height = 28
r += 2

_no_sale_names = '、'.join(sorted(x['name'] for x in rows if abs(x['q10']) <= 1e-9 and x['q0'] > 0)) or '（无）'
if ADD_BACK:
    sec(ws, r, 1, 10, f"二、{NEXT_FULL}动销（回加至库存，不参与 {_d_end.month} 月结算，共 {T['q11']} 件）"); r += 1
    hdr(ws, r, 1, ["商品条码", "商品名称", "规格", "单位", f"{NEXT_LABEL}销量（件）", "", "", "", "", "占比"], 30, fl=BLUE)
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=9)
    r += 1
    nov_bars = sorted([b for b in nov if nov[b] > 0], key=lambda b: -nov[b])
    NOV_TOTAL = sum(nov.values())
    for i, b in enumerate(nov_bars, 1):
        bg = WHITE if i % 2 else GREY
        cell = ws.cell(row=r, column=1); cell.value = b; cell.font = F(9); cell.alignment = A("center")
        cell.number_format = "@"; cell.fill = fill(bg)
        mput(ws, r, 2, r, 2, oct_name.get(b, ""), F(9), A("left", "center", True, 0), bg)
        mput(ws, r, 3, r, 3, oct_spec.get(b, ""), F(9), A("center"), bg)
        mput(ws, r, 4, r, 4, oct_unit.get(b, ""), F(9), A("center"), bg)
        cell = ws.cell(row=r, column=5); num(cell, nov[b]); cell.fill = fill(bg)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=9)
        cell = ws.cell(row=r, column=10); cell.value = nov[b] / NOV_TOTAL
        cell.number_format = "0.0%"; cell.font = F(9); cell.alignment = A("center"); cell.fill = fill(bg)
        ws.row_dimensions[r].height = 24
        r += 1
    mput(ws, r, 1, r, 4, "合　计", F(9.5, True, WHITE), A("center"), NAVY)
    cell = ws.cell(row=r, column=5); num(cell, NOV_TOTAL); cell.font = F(10, True, WHITE)
    cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=9)
    cell = ws.cell(row=r, column=10); cell.value = 1.0; cell.number_format = "0.0%"
    cell.font = F(9.5, True, WHITE); cell.alignment = A("center")
    cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
    ws.row_dimensions[r].height = 28
    r += 2
else:
    r += 1
mput(ws, r, 1, r, 10,
     (f"注：门店分布仅统计本期动销；导出日前销售（{T['q11']} 件）仅列合计。当期无动销单品：{_no_sale_names}。销售流水有记录日期 {COV_TXT}。"
      if ADD_BACK else
      f"注：门店分布仅统计本期动销。当期无动销单品：{_no_sale_names}。销售流水有记录日期 {COV_TXT}。"),
     F(8, False, DGREY), A("left", "center", True, 0))
ws.row_dimensions[r].height = 26
print("05 ok")

# ==================== 06 乙方库存明细 ====================
ws = out.create_sheet("06-乙方库存明细")
setup(ws, [4, 32, 14, 9, 11, 8, 8, 5, 11, 9, 10, 16], "landscape")
ws.row_dimensions[1].height = 10
mput(ws, 2, 1, 2, 12, f"附 表 06　乙方 {_d_asof:%Y%m%d} 库存明细", F(15, True, NAVY), A("center"))
ws.row_dimensions[2].height = 32
mput(ws, 3, 1, 3, 12, f"数据来源：乙方提供的 {ASOF} 库存数据　·　库存时点 {ASOF}　·　金额单位：人民币元",
     F(8.5, False, DGREY), A("center"))
ws.row_dimensions[3].height = 20
ws.row_dimensions[4].height = 6
r = 5
hdr(ws, r, 1, ["序号", "商品名称", "商品条码", "规格", "商品分类\n（乙方系统）", "库存数量", "零售价",
               "品牌", "结算属性", "期末合同参考价", "期末参考货值", "备　注"], 34)
r += 1
settle_by_bar = {x["bar"]: x for x in rows if x["bar"]}
settle_by_n = {norm(x["name"]): x for x in rows}
inv_sorted = sorted(inv, key=lambda x: (0 if norm(x["name"]) in settle_by_n else 1, x["name"]))
tot_q, tot_v = 0, 0.0
for i, x in enumerate(inv_sorted, 1):
    st = settle_by_bar.get(x["bar"]) or settle_by_n.get(norm(x["name"]))
    bg = WHITE if i % 2 else GREY
    mput(ws, r, 1, r, 1, i, F(9), A("center"), bg)
    mput(ws, r, 2, r, 2, x["name"], F(9), A("left", "center", True, 0), bg)
    cell = ws.cell(row=r, column=3); cell.value = x["bar"]; cell.font = F(9); cell.alignment = A("center")
    cell.number_format = "@"; cell.fill = fill(bg)
    _ms6 = MS_BAR.get(x["bar"], {})
    mput(ws, r, 4, r, 4, (_ms6.get("spec") or x["spec"]), F(9), A("center"), bg)
    mput(ws, r, 5, r, 5, x["cat"], F(9), A("center"), bg)
    cell = ws.cell(row=r, column=6); num(cell, x["qty"]); cell.fill = fill(bg)
    cell = ws.cell(row=r, column=7); num(cell, x["retail"], "#,##0.00"); cell.fill = fill(bg)
    mput(ws, r, 8, r, 8, x.get("brand") or "—", F(9), A("center"), bg)
    mput(ws, r, 9, r, 9, "结算类" if st else "非结算类", F(8.5, True, BLUE if st else "7F7F7F"), A("center"), VLBLUE if st else GREY)
    cell = ws.cell(row=r, column=10)
    if st:
        num(cell, st["price"], "#,##0.00"); cell.fill = fill(bg)
        cell = ws.cell(row=r, column=11); num(cell, float(money(x["qty"] * st["price"])), "#,##0.00"); cell.fill = fill(bg)
        tot_v += float(money(x["qty"] * st["price"]))
    else:
        cell.value = "—"; cell.font = F(9, c="A6A6A6"); cell.alignment = A("center"); cell.fill = fill(bg)
        cell = ws.cell(row=r, column=11); cell.value = "—"; cell.font = F(9, c="A6A6A6")
        cell.alignment = A("center"); cell.fill = fill(bg)
    note = ""
    if not st: note = "未列入甲方挂账台账，不属甲方铺货货品；归属另行确定，数量以乙方库存为准"
    if abs(x["qty"]) <= 1e-9: note = "本期无库存"
    mput(ws, r, 12, r, 12, note or "—", F(8, c=DGREY), A("left", "center", True, 1), bg)
    ws.row_dimensions[r].height = 24
    tot_q += x["qty"]
    r += 1
mput(ws, r, 1, r, 5, f"结算类小计（{len(inv) - len(extra)} 个单品）", F(9.5, True, WHITE), A("center"), NAVY)
cell = ws.cell(row=r, column=6); num(cell, tot_q - EXTRA_Q); cell.font = F(10, True, WHITE)
cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
for j in (7, 8, 9, 10):
    cell = ws.cell(row=r, column=j); cell.value = None; cell.fill = fill(NAVY)
    cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
cell = ws.cell(row=r, column=11); num(cell, float(money(tot_v)), "#,##0.00"); cell.font = F(10, True, WHITE)
cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
mput(ws, r, 12, r, 12, "按期末合同价参考估值", F(8, False, WHITE), A("center"), NAVY)
ws.row_dimensions[r].height = 26
r += 1
mput(ws, r, 1, r, 5, f"非结算类小计（{len(extra)} 个单品）", F(9.5, True, WHITE), A("center"), "7F7F7F")
cell = ws.cell(row=r, column=6); num(cell, EXTRA_Q); cell.font = F(10, True, WHITE)
cell.fill = fill("7F7F7F"); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
for j in (7, 8, 9, 10):
    cell = ws.cell(row=r, column=j); cell.value = None; cell.fill = fill("7F7F7F")
    cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
cell = ws.cell(row=r, column=11); cell.value = "不估值"; cell.font = F(9, False, WHITE)
cell.alignment = A("center"); cell.fill = fill("7F7F7F")
cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
mput(ws, r, 12, r, 12, "未列入甲方挂账台账", F(8, False, WHITE), A("center"), "7F7F7F")
ws.row_dimensions[r].height = 26
r += 2
mput(ws, r, 1, r, 12, f"说明：{len(inv)} 个单品中 {len(inv) - len(extra)} 个为结算类（已挂账），期末参考货值 {tot_v:,.2f} 元，"
     f"按期末合同价估值，不参与货款结算或差异倒挤；{len(extra)} 个单品共 {EXTRA_Q:,} 件未在甲方挂账，不作估值。",
     F(8, False, DGREY), A("left", "center", True, 0))
ws.row_dimensions[r].height = 30
ws.freeze_panes = "C6"
print("06 ok")

# ==================== 07 乙方销售流水 ====================
ws = out.create_sheet("07-乙方销售流水")
setup(ws, [18, 14, 12, 28, 14, 9, 5, 8, 8, 9, 18, 13, 14, 28], "landscape")
ws.row_dimensions[1].height = 10
mput(ws, 2, 1, 2, 14, f"附 表 07　乙方 {PERIOD} 销售／退货流水与取价明细", F(15, True, NAVY), A("center"))
ws.row_dimensions[2].height = 32
mput(ws, 3, 1, 3, 14, f"数据来源：乙方 {_d_end.year} 年 {_d_end.month} 月流水＋合同价格政策　·　共 {len(detail)} 行 / 净数量 {sum(x[7] for x in detail)} 件　·　货款 {T['settle']:,.2f} 元",
     F(8.5, False, DGREY), A("center"))
ws.row_dimensions[3].height = 20
ws.row_dimensions[4].height = 6
r = 5
hdr(ws, r, 1, ["流水号", "业务日期", "门店", "商品名称", "商品条码", "规格", "单位", "数量\n（退货为负）", "品牌",
               "业务类型", "原销售流水号", "合同价／原销售价", "货款金额\n（元）", "取价依据"], 32)
r += 1
for i, d in enumerate(detail, 1):
    bg = WHITE if i % 2 else GREY
    cell = ws.cell(row=r, column=1); cell.value = str(d[0]); cell.font = F(8.5); cell.alignment = A("center")
    cell.number_format = "@"; cell.fill = fill(bg)
    _raw = str(d[1]).strip()
    for _fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
        try:
            _dt = _datetime.strptime(_raw, _fmt); break
        except ValueError:
            _dt = _raw
    _c = mput(ws, r, 2, r, 2, _dt, F(8.5, c=DGREY), A("center"), bg)
    if isinstance(_dt, _datetime):
        _c.number_format = "yyyy-mm-dd hh:mm"
    mput(ws, r, 3, r, 3, d[2], F(8.5), A("center"), bg)
    mput(ws, r, 4, r, 4, d[3], F(8.5), A("left", "center", True, 0), bg)
    cell = ws.cell(row=r, column=5); cell.value = str(d[4]); cell.font = F(8.5); cell.alignment = A("center")
    cell.number_format = "@"; cell.fill = fill(bg)
    mput(ws, r, 6, r, 6, d[5], F(8.5), A("center"), bg)
    mput(ws, r, 7, r, 7, d[6], F(8.5), A("center"), bg)
    cell = ws.cell(row=r, column=8); num(cell, d[7]); cell.fill = fill(bg)
    mput(ws, r, 9, r, 9, d[8], F(8.5), A("center"), bg)
    mput(ws, r, 10, r, 10, d[9], F(8.5), A("center"), bg)
    mput(ws, r, 11, r, 11, d[10] or "—", F(8.5), A("center"), bg)
    cell = ws.cell(row=r, column=12)
    if d[11] is None:
        cell.value = "—"; cell.font = F(8.5, c=DGREY); cell.alignment = A("center")
    else:
        num(cell, d[11], "#,##0.00")
    cell.fill = fill(bg)
    cell = ws.cell(row=r, column=13)
    if d[12] is None:
        cell.value = "—"; cell.font = F(8.5, c=DGREY); cell.alignment = A("center")
    else:
        num(cell, d[12], "#,##0.00")
    cell.fill = fill(bg)
    mput(ws, r, 14, r, 14, d[13] or "非结算类，不计货款", F(7.5, c=DGREY), A("left", "center", True, 1), bg)
    ws.row_dimensions[r].height = 24
    r += 1
mput(ws, r, 1, r, 7, "合　计", F(9.5, True, WHITE), A("center"), NAVY)
cell = ws.cell(row=r, column=8); num(cell, sum(x[7] for x in detail)); cell.font = F(10, True, WHITE)
cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
for _col in (9, 10, 11, 12, 14):
    cell = ws.cell(row=r, column=_col); cell.value = None; cell.fill = fill(NAVY)
    cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
cell = ws.cell(row=r, column=13); num(cell, T["settle"], "#,##0.00"); cell.font = F(10, True, WHITE)
cell.fill = fill(NAVY); cell.border = Border(left=THIN, right=THIN, top=MED, bottom=MED)
ws.row_dimensions[r].height = 28
r += 2
_grp = Counter((str(d[0]).strip(), str(d[4]).strip(), d[7]) for d in detail)
_split = sorted({g[0] for g, n in _grp.items() if n > 1})
if _split:
    _shown = "、".join(_split[:4]) + ("…" if len(_split) > 4 else "")
    _note = (f"核对说明：存在 {len(_split)} 组“同一流水号 + 同一条码 + 同一数量”的拆行记录"
             f"（流水号 {_shown}），已逐行计入合计数、未作去重或调整；是否属正常拆行，请双方核对。")
else:
    _note = "核对说明：流水明细未发现重复拆行；销售按各自销售日合同价计算，退货按原销售价冲减，金额不取绝对值。"
mput(ws, r, 1, r, 14, _note, F(8, False, DGREY), A("left", "center", True, 0))
ws.row_dimensions[r].height = 34
ws.freeze_panes = "A6"
print("07 ok")

# ==================== 08 效期预警 ====================
ws = out.create_sheet("08-效期预警")
setup(ws, [6, 34, 11, 14, 12, 10, 9, 9, 9, 11, 14, 14], "landscape")
ws.row_dimensions[1].height = 10
mput(ws, 2, 1, 2, 12, "附 表 08　批次效期预警", F(15, True, NAVY), A("center"))
ws.row_dimensions[2].height = 32
mput(ws, 3, 1, 3, 12,
     f"基准日：{P_END}（对账截止日）　·　数据来源：{BATCH_SRC or '未提供'}　·　"
     f"期末结存＝期初批次数量 − 本期动销（按有效期升序 FIFO 扣减）",
     F(8.5, False, DGREY), A("center"))
ws.row_dimensions[3].height = 20
ws.row_dimensions[4].height = 6
r = 5

if not BATCH:
    sec(ws, r, 1, 12, "一、本期未提供批次效期数据"); r += 1
    mput(ws, r, 1, r, 12,
         BATCH_NOTE + "。请在甲方挂账表增加「批次号／有效期至」两列，或另提供批次效期表后重跑；"
         "本表不推测效期分布，也不按品名或经验赋值。",
         F(9), A("left", "center", True, 1), LAMBER)
    ws.row_dimensions[r].height = 44; r += 2
    sec(ws, r, 1, 12, "二、需要补充的材料（任一即可）"); r += 1
    for _t in ["① 甲方挂账表增加「批次号」「有效期至」列——推荐，与挂账数据同源，无需额外文件；",
               "② 或提供独立批次效期表：条码、品名、批次号、生产日期、有效期至、数量；",
               "③ 若只有生产日期与保质期，脚本按「生产日期＋保质期（月）」推算到期日；",
               "④ 三者都没有时，本表只出本说明页，并在 00 表列入待确认事项。"]:
        mput(ws, r, 1, r, 12, _t, F(9), A("left", "center", True, 1))
        ws.row_dimensions[r].height = 24; r += 1
else:
    sec(ws, r, 1, 12, "一、预警汇总（按期末结存的剩余天数分档）"); r += 1
    hdr(ws, r, 1, ["预警档", "", "", "批次数", "SKU 数", "件数", "", "", "", "", "", ""], 26)
    for _c1, _c2 in ((1, 3), (6, 12)):
        ws.merge_cells(start_row=r, start_column=_c1, end_row=r, end_column=_c2)
    r += 1
    for _t in EXP_TIERS:
        _sm = EXP_SUM[_t]
        _fl = (LRED if _t == "已过期" else LAMBER if _t in ("30 天内", "90 天内")
               else GREY if _t == "180 天内" else WHITE)
        mput(ws, r, 1, r, 3, _t, F(9, True), A("center"), _fl)
        for _c, _v in ((4, _sm["lots"]), (5, len(_sm["skus"])), (6, _sm["qty"])):
            _cell = ws.cell(row=r, column=_c); num(_cell, _v); _cell.fill = fill(_fl)
        for _c in range(7, 13):
            _cell = ws.cell(row=r, column=_c); _cell.value = None; _cell.fill = fill(_fl)
            _cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        ws.row_dimensions[r].height = 22; r += 1
    mput(ws, r, 1, r, 12,
         f"合计：期末结存 {len([_b for _b in BATCH if _b['left'] > 0])} 个批次 / {EXP_TOTAL_LEFT:,} 件。"
         f"其中 90 天内到期或已过期 {EXP_ALERT_SKU} 个 SKU、{EXP_ALERT_QTY:,} 件（含 30 天内与已过期）。",
         F(9, True, NAVY), A("left", "center", True, 1), LBLUE)
    ws.row_dimensions[r].height = 24; r += 2

    sec(ws, r, 1, 12, "二、预警明细（仅列期末有结存的批次，按剩余天数升序）"); r += 1
    hdr(ws, r, 1, ["序号", "商品名称", "规格", "批次号", "有效期至", "剩余天数",
                   "期初数量", "本期扣减", "期末结存", "预警档", "商品条码", "备注"], 30)
    r += 1
    _spec_of = {b: (v.get("spec") or "") for b, v in MS_BAR.items()}
    _fifo_lots = {(n, l) for n, ls, _ in FIFO_ISSUES for l in ls} if FIFO_ISSUES else set()
    _live = [_b for _b in BATCH if _b["left"] > 0]
    for _i, _b in enumerate(sorted(_live, key=lambda z: (z["days"], z["name"])), 1):
        _bg = WHITE if _i % 2 else GREY
        mput(ws, r, 1, r, 1, _i, F(8.5), A("center"), _bg)
        mput(ws, r, 2, r, 2, _b["name"], F(8.5), A("left", "center", True, 0), _bg)
        mput(ws, r, 3, r, 3, _spec_of.get(_b["bar"], "") or "—", F(8.5), A("center"), _bg)
        mput(ws, r, 4, r, 4, _b["lot"], F(8.5), A("center"), _bg)
        _c = mput(ws, r, 5, r, 5, _b["exp"], F(8.5), A("center"), _bg)
        _c.number_format = "yyyy-mm-dd"
        _cell = ws.cell(row=r, column=6); num(_cell, _b["days"]); _cell.fill = fill(_bg)
        if _b["days"] < 0:
            _cell.font = F(8.5, True, RED)
        for _c, _v in ((7, _b["qty"]), (8, _b.get("used", 0)), (9, _b["left"])):
            _cell = ws.cell(row=r, column=_c); num(_cell, _v); _cell.fill = fill(_bg)
        _tf = {"已过期": LRED, "30 天内": LAMBER, "90 天内": LAMBER,
               "180 天内": GREY, "正常（>180 天）": WHITE}[_b["tier"]]
        mput(ws, r, 10, r, 10, _b["tier"], F(8.5, True), A("center"), _tf)
        _cell = ws.cell(row=r, column=11); _cell.value = _b["bar"]; _cell.font = F(8.5)
        _cell.alignment = A("center"); _cell.number_format = "@"; _cell.fill = fill(_bg)
        _note = "后进先出迹象" if (_b["name"], _b["lot"]) in _fifo_lots else "—"
        mput(ws, r, 12, r, 12, _note, F(8, c=RED if _note != "—" else DGREY),
             A("left", "center", True, 0), _bg)
        ws.row_dimensions[r].height = 19; r += 1
    r += 1

    sec(ws, r, 1, 12, "三、批次结存与账面勾稽／FIFO 校验"); r += 1
    if EXP_RECON:
        for _bar_, _l_, _t_, _why in EXP_RECON[:20]:
            mput(ws, r, 1, r, 12,
                 f"⚠ 条码 {_bar_ or '（无条码）'}：期末批次结存 {_l_:,} 件，账面期末结存 "
                 f"{'（无对应挂账行）' if _t_ is None else format(_t_, ',') + ' 件'}；{_why}",
                 F(8.5), A("left", "center", True, 1), LAMBER)
            ws.row_dimensions[r].height = 20; r += 1
        mput(ws, r, 1, r, 12,
             f"说明：共 {len(EXP_RECON)} 个条码的批次结存与账面期末结存不一致（合计差 {_exp_diff_qty:,} 件），"
             f"多为批次数量与挂账数量口径不同或批次缺失；请核对后重跑，本表不做跨批次推测。",
             F(8.5, False, DGREY), A("left", "center", True, 1)); ws.row_dimensions[r].height = 32; r += 1
    else:
        mput(ws, r, 1, r, 12,
             f"批次结存与账面逐条码一致：期末批次结存合计 {EXP_TOTAL_LEFT:,} 件 = 账面期末理论结存。",
             F(9), A("left", "center", True, 1), LGREEN)
        ws.row_dimensions[r].height = 20; r += 1
    if FIFO_ISSUES:
        for _n, _earlier, _later in FIFO_ISSUES[:12]:
            mput(ws, r, 1, r, 12,
                 f"⚠ 「{_n}」较早批次（{'、'.join(_earlier[:3])}）仍有余额，但较晚批次"
                 f"（{'、'.join(_later[:3])}）已发生扣减，与先进先出原则不符，请确认批次核销口径。",
                 F(8.5), A("left", "center", True, 1), LAMBER)
            ws.row_dimensions[r].height = 26; r += 1
    else:
        mput(ws, r, 1, r, 12, "FIFO 校验：未发现后进先出迹象（后批次被扣减前，更早批次均已清零）。",
             F(9), A("left", "center", True, 1), LGREEN)
        ws.row_dimensions[r].height = 20; r += 1
    r += 1

    sec(ws, r, 1, 12, "四、口径说明"); r += 1
    for _t in [
        "1. 数量口径：期初数量取甲方挂账表各批次的「未结算数量」；期末结存＝期初数量 − 本期动销，按有效期升序 FIFO 扣减。",
        "2. 分档口径：按期末结存距对账截止日的剩余天数分档（已过期 / 30 天内 / 90 天内 / 180 天内 / 正常）。",
        "3. FIFO 口径：先进先出为默认原则；若挂账表显示较晚批次已扣减、更早批次仍有余额，判为后进先出迹象并单列。",
        "4. 本表为效期提醒，供乙方优先动销或申请退回；具体处理方式与费用承担由双方书面确认，本表不作归因。",
    ]:
        mput(ws, r, 1, r, 12, _t, F(8.5), A("left", "center", True, 1))
        ws.row_dimensions[r].height = _fit_h(_t); r += 1
ws.freeze_panes = "A6"
print("08 ok")

# ==================== 09 试用装与配赠 ====================
ws = out.create_sheet("09-试用装与配赠")
setup(ws, [6, 34, 11, 11, 13, 15, 12, 26], "landscape")
ws.row_dimensions[1].height = 10
mput(ws, 2, 1, 2, 8, "附 表 09　试用装 / 中小样 / 配赠 使用情况", F(15, True, NAVY), A("center"))
ws.row_dimensions[2].height = 32
mput(ws, 3, 1, 3, 8,
     f"甲方数据：挂账表「物料属性」列（正装以外的促销物料）　·　乙方数据：{PROMO_USE_SRC or '未提供'}",
     F(8.5, False, DGREY), A("center"))
ws.row_dimensions[3].height = 20
ws.row_dimensions[4].height = 6
r = 5
_PROMO_QTY = sum(_p["out_qty"] for _p in PROMO.values())
_BASIS = str(CFG.get("promo_basis") or "outbound").lower()

sec(ws, r, 1, 8, "一、口径说明"); r += 1
for _t in [
    "1. 试用装 / 中小样 / 配赠由甲方调拨出库，乙方是否入库取决于终端渠道规则，故不参与货款结算与库存数量勾稽。",
    "2. 本表只列示数量，不计算金额（单价与费用口径由财务另行确认）；出库数量取甲方挂账表对应物料的「未结算数量」。",
    ("3. 乙方已提供使用数据：" + PROMO_USE_SRC + "；差异＝出库－使用。"
     if (PROMO and PROMO_USE_SRC) else
     "3. 乙方本期未提供使用数据。按「已出库即计列」口径列示，标注未经乙方确认；乙方补报后可重算。"
     if PROMO else
     "3. 本期无促销物料出库，不涉及使用数据。"),
    "4. 本表不作归因，差异处理与费用承担由双方书面确认。",
]:
    mput(ws, r, 1, r, 8, _t, F(8.5), A("left", "center", True, 1))
    ws.row_dimensions[r].height = _fit_h(_t); r += 1
r += 1

if not PROMO:
    sec(ws, r, 1, 8, "二、本期无促销物料出库"); r += 1
    mput(ws, r, 1, r, 8,
         "本期甲方挂账表中未出现「物料属性」为试用装／中小样／配赠的行。若确有调拨出库，"
         "请在该表补「物料属性」列后重跑；本表不按品名关键词猜测物料属性。",
         F(9), A("left", "center", True, 1), LAMBER)
    ws.row_dimensions[r].height = 36; r += 1
else:
    sec(ws, r, 1, 8, "二、按物料属性汇总"); r += 1
    hdr(ws, r, 1, ["物料属性", "", "SKU 数", "出库件数", "乙方报告使用", "差异", "", ""], 26)
    for _c1, _c2 in ((1, 2), (6, 8)):
        ws.merge_cells(start_row=r, start_column=_c1, end_row=r, end_column=_c2)
    r += 1
    for _k in ("试用装", "中小样", "配赠"):
        _items = [x for x in PROMO.values() if x["kind"] == _k]
        if not _items:
            continue
        _oq = sum(x["out_qty"] for x in _items)
        _uses = []
        for x in _items:
            _u = PROMO_USE.get(x["bar"]) if x["bar"] else None
            if _u is None:
                _u = PROMO_USE.get(norm(x["name"]))
            if _u is not None:
                _uses.append(_u)
        _have = len(_uses) == len(_items)
        mput(ws, r, 1, r, 2, _k, F(9, True), A("center"), VLBLUE)
        _cell = ws.cell(row=r, column=3); num(_cell, len(_items))
        _cell = ws.cell(row=r, column=4); num(_cell, _oq)
        _cell = ws.cell(row=r, column=5)
        if _have:
            num(_cell, sum(_uses))
            _cell = ws.cell(row=r, column=6); num(_cell, _oq - sum(_uses))
        else:
            _cell.value = "—"; _cell.alignment = A("center"); _cell.font = F(9, c=DGREY)
            _cell = ws.cell(row=r, column=6); _cell.value = "—"
            _cell.alignment = A("center"); _cell.font = F(9, c=DGREY)
        for _c in range(7, 9):
            _cell = ws.cell(row=r, column=_c); _cell.value = None
        ws.row_dimensions[r].height = 22; r += 1
    mput(ws, r, 1, r, 8,
         f"合计：{len(PROMO)} 个促销物料、出库 {_PROMO_QTY:,} 件"
         + ("；乙方已提供使用数据，差异见上表。" if PROMO_USE_SRC else "；乙方未提供使用数据，差异不计算。"),
         F(9, True, NAVY), A("left", "center", True, 1), LBLUE)
    ws.row_dimensions[r].height = 24; r += 2

    sec(ws, r, 1, 8, "三、逐 SKU 明细"); r += 1
    hdr(ws, r, 1, ["序号", "商品名称", "规格", "物料属性", "出库数量", "乙方报告使用", "差异", "商品条码"], 30)
    r += 1
    for _i, _x in enumerate(sorted(PROMO.values(), key=lambda z: (z["kind"], z["name"])), 1):
        _bg = WHITE if _i % 2 else GREY
        _u = PROMO_USE.get(_x["bar"]) if _x["bar"] else None
        if _u is None:
            _u = PROMO_USE.get(norm(_x["name"]))
        mput(ws, r, 1, r, 1, _i, F(8.5), A("center"), _bg)
        mput(ws, r, 2, r, 2, _x["name"], F(8.5), A("left", "center", True, 0), _bg)
        mput(ws, r, 3, r, 3, _spec_of.get(_x["bar"], "") or "—", F(8.5), A("center"), _bg)
        mput(ws, r, 4, r, 4, _x["kind"], F(8.5, True), A("center"), _bg)
        _cell = ws.cell(row=r, column=5); num(_cell, _x["out_qty"]); _cell.fill = fill(_bg)
        _cell = ws.cell(row=r, column=6)
        if _u is None:
            _cell.value = "—"; _cell.alignment = A("center"); _cell.font = F(8.5, c=DGREY)
            _cell.fill = fill(_bg)
            _cell = ws.cell(row=r, column=7); _cell.value = "—"
            _cell.alignment = A("center"); _cell.font = F(8.5, c=DGREY); _cell.fill = fill(_bg)
        else:
            num(_cell, _u); _cell.fill = fill(_bg)
            _cell = ws.cell(row=r, column=7); num(_cell, _x["out_qty"] - _u); _cell.fill = fill(_bg)
            if _x["out_qty"] != _u:
                _cell.font = F(8.5, True, AMBER)
        _cell = ws.cell(row=r, column=8); _cell.value = _x["bar"]; _cell.font = F(8.5)
        _cell.alignment = A("center"); _cell.number_format = "@"; _cell.fill = fill(_bg)
        ws.row_dimensions[r].height = 20; r += 1
    r += 1

    sec(ws, r, 1, 8, "四、费用核算参考与待确认"); r += 1
    _bl = ("本表按甲方出库数量计列（未经乙方确认），供费用核算参考；最终结算金额与开票口径由双方书面确认。"
           if not PROMO_USE_SRC and _BASIS != "none" else
           "本表仅列示甲方出库数量，不认定乙方使用量；待乙方补报后重算。")
    mput(ws, r, 1, r, 8, _bl, F(9), A("left", "center", True, 1), LAMBER)
    ws.row_dimensions[r].height = _fit_h(_bl); r += 1
ws.freeze_panes = "A6"
print("09 ok")

# ==================== v1.7 附表数据交接 ====================
if V17:
    def _d2s(v):
        return v.isoformat() if hasattr(v, "isoformat") else v
    _sc = {
        "period": PERIOD, "asof": str(_PD.get("asof") or ASOF),
        "end": P_END, "period_cn": P_CN,
        "party_a": PA_NAME, "party_b": PB_NAME,
        "expiry": KJD.get("expiry") or {},
        "sheet_map": SHEET_MAP,
        "ledger": LEDGER,
        "batch": [{"name": b.get("name", ""), "bar": b.get("bar", ""),
                   "lot": b.get("lot", ""), "exp": _d2s(b.get("exp")),
                   "qty": b.get("qty", 0), "init": b.get("init", 0),
                   "left": b.get("left", b.get("qty", 0)),
                   "price": b.get("price", 0)} for b in BATCH],
        "promo": [{"name": x.get("name", ""), "kind": x.get("kind", ""),
                   "bar": x.get("bar", ""), "out_qty": x.get("out_qty", 0)}
                  for x in PROMO.values()],
        "rows": [{"bar": r.get("bar", ""), "name": r.get("name", ""),
                  "q0": r.get("q0", 0), "q10": r.get("q10", 0),
                  "theo": r.get("theo", 0)} for r in rows],
        "totals": {k: v for k, v in T.items()},
        "rule_result": jsonable(RULE_RESULT),
        "batch_src": BATCH_SRC,
        "out": OUT,
        "cfg_v17": CFG.get("v17") or {},
    }
    _sc_path = os.path.join(BASE, "_v17_data.json")
    with open(_sc_path, "w", encoding="utf-8") as _fh:
        json.dump(_sc, _fh, ensure_ascii=False, indent=1, default=str)
    print(f"v17 sidecar: {_sc_path}", file=sys.stderr)

# ==================== 保存 ====================
for sn in ["00-对账说明", "01-对账单", "02-库存差异表"]:
    out[sn].sheet_properties.tabColor = NAVY
out["02-库存差异表"].sheet_properties.tabColor = RED
for sn in ["03-源数据对照", "04-甲方挂账明细", "05-乙方销售汇总", "06-乙方库存明细",
           "07-乙方销售流水", "08-效期预警", "09-试用装与配赠"]:
    out[sn].sheet_properties.tabColor = "A6A6A6"
out["08-效期预警"].sheet_properties.tabColor = AMBER
out["09-试用装与配赠"].sheet_properties.tabColor = BLUE
TITLES = {"01-对账单": "12:12", "02-库存差异表": "7:7", "03-源数据对照": "5:5",
          "04-甲方挂账明细": "5:5", "05-乙方销售汇总": "6:6",
          "06-乙方库存明细": "5:5", "07-乙方销售流水": "5:5",
          "08-效期预警": "5:5", "09-试用装与配赠": "5:5"}
for sn in out.sheetnames:
    w = out[sn]
    w.print_area = f"A1:{get_column_letter(w.max_column)}{w.max_row}"
    if sn in TITLES:
        w.print_title_rows = TITLES[sn]
    w.oddHeader.center.text = f"{PA_NAME} × {PB_NAME}　·　{P_CN}对账"
    w.oddHeader.center.size = 8
    w.oddHeader.center.color = "808080"
    w.oddFooter.center.text = "第 &P 页 / 共 &N 页"
    w.oddFooter.center.size = 8
    w.oddFooter.center.color = "808080"
    w.oddFooter.right.text = f"{PA_NAME}  ·  {P_CN}对账"
    w.oddFooter.right.size = 8
    w.oddFooter.right.color = "808080"
# 说明：报表内所有主体/期间/条款文案均在上方按 CFG 直接生成；
# 早期版本曾用「保存前对整簿做字符串替换」兜底，会误伤数据区且漏替条款，已移除。
out.active = 0
# 注意：失败时 stdout 必须保持 0 字节，所有 stdout 输出都放到保存成功之后
_tmp_out = OUT + ".tmp.xlsx"
try:
    out.save(_tmp_out)
    os.replace(_tmp_out, OUT)
except PermissionError as _e:
    if os.path.exists(_tmp_out):
        os.remove(_tmp_out)
    raise SystemExit(f"\n[写入失败] {OUT}\n文件可能正被 Excel 打开，请关闭后重试。原始错误：{_e}")
except Exception as _e:
    if os.path.exists(_tmp_out):
        os.remove(_tmp_out)
    raise SystemExit(f"\n[写入失败] {OUT}\n磁盘空间不足或路径不可写，原文件未受影响。原始错误：{_e}")
if V17:
    sys.path.insert(0, _SK)
    from check_public_copy import clean_properties
    clean_properties(OUT)

# 无论是否启用 V17，均输出可机读规则结果与 writeback_ledger.py 可直接消费的回写负载。
def _write_json_atomic(path, payload):
    _tmp = path + ".tmp"
    with open(_tmp, "w", encoding="utf-8") as _fh:
        json.dump(payload, _fh, ensure_ascii=False, indent=2)
        _fh.write("\n")
    os.replace(_tmp, path)


try:  # 与主档取列同款健壮导入：哈希统一走 schema_common 的 1MiB 分块实现（结果与整读逐字节一致）
    from schema_common import sha256_file as _sha256_file
except Exception as _sha_import_error:  # noqa: BLE001 - 依赖缺失时中文中止
    _bail(f"输出凭证哈希需要 schema_common.py（{_sha_import_error}）")

_workbook_sha = _sha256_file(OUT)
_WRITEBACK = {
    "period": PERIOD, "asof": ASOF,
    "contract_no": CONTRACT_NO, "business_type": "铺货代销", "statement_no": STATEMENT_NO,
    "opening_qty": T["q0"], "opening_amount": OPEN_VAL,
    "settle_qty": T["q10"], "settle_amount": SETTLEMENT_TOTAL,
    "diff_qty": T["diff"], "diff_amount": T["diff_amt"],
    "sticker_amount": TH_TOTAL, "sticker_pay": "待定", "status": "待确认",
    "note": ("铺货代销；期初金额为期末合同价参考估值；盈亏按期末合同价列示，"
             + ("经本期单次授权纳入结算总额。" if VARIANCE_INCLUDED else "未纳入本期结算总额，纳入须逐期单次授权。")),
}
_RESULT_PATH = os.path.join(BASE, "_结算结果.json")
_WRITEBACK_PATH = os.path.join(BASE, "_回写.json")
_RESULT = {
    "schema": "dealer-ops/settlement-result@1",
    "business_type": BUSINESS_TYPE,
    "workbook": {"path": os.path.abspath(OUT), "sha256": _workbook_sha},
    "rule_result": jsonable(RULE_RESULT),
    "writeback": _WRITEBACK,
    "writeback_path": _WRITEBACK_PATH,
    "master": {"path": os.path.abspath(_MASTER), "sha256": _sha256_file(_MASTER)},
    "sources": {
        role: {"path": os.path.abspath(path), "sha256": _sha256_file(path)}
        for role, path in (("jia", _P_JIA), ("sales", _P_OCT), ("inventory", _P_INV))
    },
}
_write_json_atomic(_WRITEBACK_PATH, _WRITEBACK)
_write_json_atomic(_RESULT_PATH, _RESULT)
emit(f"{_OUT_LABEL} {OUT}")
emit(f"SAVED: {OUT}")
emit(f"RESULT: {_RESULT_PATH}")
emit(f"WRITEBACK: {_WRITEBACK_PATH}")
emit("大小 %.1f KB" % (os.path.getsize(OUT) / 1024))
