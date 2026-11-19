#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_plan.py —— 把「用户确认的值」编译成可执行的填充计划（只读模板，不写文件）

做四件事：
  1. 逐字段校验（信用代码校验位、日期逻辑、手机号、金额大小写、百分比区间 …）
  2. 跨字段一致性校验（起止日期、保证金大小写、返利高低、附件日期与主合同一致）
  3. 按模板现状解析每个待填项的精确落点，渲染最终写入文本
  4. 产出 plan.json（供 apply_fill.py 执行）与可直接贴给用户的对照表

用法：
  python3 build_plan.py --spec spec/jx.json --doc 模板.docx --values values.json --out plan.json

退出码：0 全部通过（生成 plan.json）；1 存在 ERROR（fail-closed：不生成可执行计划，
     只留 <out>.rejected 供排查，并删除上一轮遗留的旧 plan.json）
计划内会写入 spec_sha256 / base_sha256 / values_digest / source_values，
验收脚本据此把「成品 == 计划」升级为「成品 == 计划 == 用户确认值」。
"""
import argparse
import hashlib
import json
import os
import re
import sys
import unicodedata
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from scan_blanks import (Document, analyse_cell, analyse_field,
                         analyse_grid, build_units)  # noqa: E402
from plan_trust import file_sha256  # noqa: E402

USCC_CHARS = "0123456789ABCDEFGHJKLMNPQRTUWXY"
USCC_W = [1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28]
CN_DIGITS = "零壹贰叁肆伍陆柒捌玖"
CN_UNITS = ["", "拾", "佰", "仟"]
CN_BIG = ["", "万", "亿", "兆"]

# 这些类型的值需要先做全角→半角归一，否则「２０」既可能被判非法，也可能被原样写进合同
NUMERIC_VTYPES = {"percent_rebate", "percent_any", "amount_num", "amount_wan",
                  "day", "int", "uscc"}

# 主合同编码规则（示例）：
#   四段式「品牌码-类型码-年份-序号」，可带一段扩展码（位于年份之前）；
#   仅大写字母、数字、半角连字符，不得含中文、空格、下划线、全角字符。
CONTRACT_TYPES = {"JX": "经销", "DX": "代销", "DL": "代理", "GY": "供应",
                  "FW": "服务", "ZX": "咨询", "XY": "意向"}
CONTRACT_NO_RE = re.compile(
    r"^([A-Z]{2})-([A-Z]{2})(?:-([A-Z]{2}))?-(\d{4})-(\d{3})$")


# ---------------------------------------------------------------- 校验工具
def check_uscc(code):
    code = (code or "").strip().upper()
    if len(code) != 18:
        return False, f"应为 18 位，当前 {len(code)} 位"
    bad = [c for c in code if c not in USCC_CHARS]
    if bad:
        return False, f"含信用代码不使用的字符：{''.join(sorted(set(bad)))}（I/O/Z/S/V 不参与编码）"
    s = sum(USCC_CHARS.index(code[i]) * USCC_W[i] for i in range(17))
    c = 31 - s % 31
    c = 0 if c == 31 else c
    if USCC_CHARS[c] != code[17]:
        return False, f"校验位不匹配，末位应为「{USCC_CHARS[c]}」"
    return True, ""


def check_contract_no(v):
    """主合同编号：格式校验（类型码/年份的一致性由 check_contract_no_consistency 管）。"""
    v = (v or "").strip()
    if v == "待定":
        # 起草阶段编号常未发放。允许占位，但交付说明里必须提示「签署前替换」。
        return True, ""
    m = CONTRACT_NO_RE.fullmatch(v)
    if not m:
        return False, ("应为「品牌码-类型码-年份-序号」（如 DM-JX-2026-001），"
                       "可带扩展段（如 DM-JX-JS-2026-001）；仅大写字母、数字、半角连字符，"
                       "不得含中文、空格、下划线、全角字符或小写字母")
    if m.group(5) == "000":
        return False, "序号自 001 起连续编排，不能为 000"
    return True, ""


def check_contract_no_consistency(code, doc_key, sign_date):
    """类型码须与合同类型一致；年份须等于签订时间所在年份。

    返回 (errors, warns)。格式本身不合法时静默跳过——那由 check_contract_no 报。
    """
    errors, warns = [], []
    code = (code or "").strip()
    if not code or code == "待定":
        return errors, warns
    m = CONTRACT_NO_RE.fullmatch(code)
    if not m:
        return errors, warns
    brand, ctype, _ext, year, _serial = m.groups()
    if brand != "DM":
        warns.append(
            f"主合同编号品牌码为「{brand}」，示例规则只登记「DM」（Demo）；"
            f"若为新增品牌请与合同管理员确认"
        )
    if ctype != doc_key:
        errors.append(
            f"主合同编号类型码为「{ctype}」，但本合同是「{doc_key}」"
            f"（JX=经销 / DX=代销），两者必须一致"
        )
    if sign_date and len(sign_date) >= 4 and year != sign_date[:4]:
        errors.append(
            f"主合同编号年份「{year}」与签订时间「{sign_date}」的年份不一致；"
            f"规则要求取合同签订年份，跨年签署也以签订日期为准"
        )
    return errors, warns


def check_date(v):
    v = (v or "").strip()
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
        return False, "格式应为 YYYY-MM-DD"
    try:
        datetime.strptime(v, "%Y-%m-%d")
    except ValueError:
        return False, "不是有效日期"
    return True, ""


def check_phone(v):
    v = (v or "").strip().replace("－", "-").replace("—", "-")
    if re.fullmatch(r"1[3-9]\d{9}", v):
        return True, ""
    if re.fullmatch(r"0\d{2,3}-?\d{7,8}(-\d{1,6})?", v):
        return True, ""
    return False, "应为 11 位手机号，或含区号的座机（如 020-88886666）"


def check_person(v):
    v = (v or "").strip()
    if not 2 <= len(v) <= 15:
        return False, "姓名长度应为 2-15 字"
    if re.search(r"[\d（）()]", v):
        return False, "姓名不应包含数字或括号"
    return True, ""


def check_company(v):
    v = (v or "").strip()
    if len(v) < 4:
        return False, "公司名称过短"
    if re.search(r"[（(]\s*[)）]", v):
        return False, "含空括号，请核实"
    if not re.search(r"(公司|企业|商行|经营部|中心|店|厂|社|事务所|合伙)$", v):
        return False, "名称未以「有限公司/商行/经营部/中心」等主体类型结尾，请核对营业执照"
    return True, ""


def check_address(v):
    v = (v or "").strip()
    if len(v) < 8:
        return False, "地址过短，应为「省+市+区+详细地址」"
    if not re.search(r"(省|市|区|县|旗|镇)", v):
        return False, "未包含行政区划（省/市/区/县）"
    if not re.search(r"(路|街|道|巷|号|村|楼|室|层|栋|园区|大厦|广场|仓)", v):
        return False, "未包含门牌/楼栋等详细指向"
    return True, ""


def check_region(v):
    v = (v or "").strip()
    if not 2 <= len(v) <= 8:
        return False, "行政区名称应为 2-8 字"
    if v.endswith(("省", "市", "区", "县")):
        return False, "模板已自带「省」「市」字样，此处不要再带后缀"
    return True, ""


def num2cn(n):
    """整数金额转中文大写（不含「人民币」「元」）

    按四位分组拼接，并处理组间缺位的「零」：
    10001 → 壹万零壹（不是「壹万壹」）；50010 → 伍万零壹拾；100000 → 壹拾万。
    """
    n = int(n)
    if n == 0:
        return "零"
    groups = []
    while n > 0:
        groups.append(n % 10000)
        n //= 10000
    out = ""
    for gi in range(len(groups) - 1, -1, -1):
        g = groups[gi]
        if g == 0:
            if out and not out.endswith("零"):
                out += "零"
            continue
        # 高位组已输出，而本组不足四位（千位为 0）时，必须补一个「零」占位
        if out and g < 1000 and not out.endswith("零"):
            out += "零"
        seg = ""
        zero = False
        for pos in range(3, -1, -1):
            d = (g // 10 ** pos) % 10
            if d == 0:
                zero = True
            else:
                if zero and seg:
                    seg += "零"
                seg += CN_DIGITS[d] + CN_UNITS[pos]
                zero = False
        out += seg + CN_BIG[gi]
    return out.rstrip("零")


def check_amount_cn(v, num=None):
    v = (v or "").strip().replace("圆", "元").replace(" ", "")
    if not v:
        return False, "不能为空"
    if not re.fullmatch(r"[零壹贰叁肆伍陆柒捌玖拾佰仟万亿元整]+", v):
        return False, "应为中文大写数字（壹贰叁…），不是阿拉伯数字"
    core = v.rstrip("元整")
    if num not in (None, ""):
        expect = num2cn(num)
        if core != expect:
            return False, f"与金额 {num} 不一致，大写应为「{expect}元」"
    return True, ""


# ---------------------------------------------------------------- 渲染
def render(tpl, value, vtype):
    if "{y}" in tpl or "{m}" in tpl or "{d}" in tpl:
        y, m, d = value.split("-")
        return tpl.replace("{y}", str(int(y))).replace("{m}", str(int(m))).replace("{d}", str(int(d)))
    return tpl.replace("{v}", str(value))


def validate_one(f, value):
    vt = f.get("vtype", "text")
    if value is None or (isinstance(value, str) and value.strip() == ""):
        return (True, "") if not f.get("required") else (False, "必填项未提供")
    value = str(value).strip()
    if "\n" in value:
        return False, "不能包含换行"
    # 控制字符与零宽字符：肉眼不可见，会污染合同文本且事后极难发现
    bad = re.search(r"[\x00-\x1f\x7f\u200b-\u200f\u202a-\u202e\u2060\ufeff]", value)
    if bad:
        return False, f"含不可见字符（U+{ord(bad.group(0)):04X}），请删掉重新粘贴"
    if vt == "contract_no":
        return check_contract_no(value)
    if vt == "date":
        return check_date(value)
    if vt == "uscc":
        return check_uscc(value)
    if vt == "phone":
        return check_phone(value)
    if vt == "person":
        return check_person(value)
    if vt == "company":
        return check_company(value)
    if vt == "address":
        return check_address(value)
    if vt == "region":
        return check_region(value)
    if vt == "text":
        return (True, "") if 1 <= len(value) <= 32 else (False, "长度应为 1-32 字")
    if vt == "channel":
        return (True, "") if 1 <= len(value) <= 20 else (False, "渠道名称建议不超过 20 字")
    if vt in ("percent_rebate", "percent_any"):
        try:
            p = float(value.rstrip("%"))
        except ValueError:
            return False, "应为数字（可带小数），如 3 或 3.5"
        if vt == "percent_rebate" and not 0 <= p <= 30:
            return False, "返利比例应在 0-30 之间"
        if vt == "percent_any" and not 0 < p <= 100:
            return False, "比例应在 0-100 之间"
        return True, ""
    if vt == "amount_num":
        return (True, "") if re.fullmatch(r"\d{1,9}", value) else (False, "应为不超过 9 位的整数（元）")
    if vt == "amount_wan":
        try:
            p = float(value)
        except ValueError:
            return False, "应为数字（单位万元）"
        return (True, "") if p > 0 else (False, "应大于 0")
    if vt == "day":
        return (True, "") if value.isdigit() and 1 <= int(value) <= 28 else \
            (False, "应为 1-28 的整数（超过 28 会碰到没有该日的月份）")
    if vt == "int":
        return (True, "") if value.isdigit() else (False, "应为非负整数")
    if vt == "amount_cn":
        return check_amount_cn(value)
    return False, f"规格里登记了未知的校验类型「{vt}」，请修正 spec 后重试"


# ---------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--spec", required=True)
    ap.add_argument("--doc", required=True)
    ap.add_argument("--values", required=True)
    ap.add_argument("--out", default="plan.json")
    args = ap.parse_args()

    spec = json.load(open(args.spec, encoding="utf-8"))
    values_in = json.load(open(args.values, encoding="utf-8"))
    values = {k: v for k, v in values_in.items() if not k.startswith("_")}

    doc = Document(args.doc)
    units = build_units(doc)

    # 已知字段白名单：拼错的键必须报错，否则会被静默丢弃（用户以为填了，其实没填）
    known_ids = {f["id"] for f in spec.get("fields", [])}
    known_ids |= {c["id"] for c in spec.get("cells", [])}
    known_ids |= {g["id"] for g in spec.get("grids", [])}
    unknown = sorted(set(values) - known_ids)
    if unknown:
        if os.path.exists(args.out):
            os.remove(args.out)
        sys.exit(
            "存在无法识别的字段 ID，已停止：\n  "
            + "\n  ".join(f"「{k}」" for k in unknown)
            + f"\n本规格（{spec['doc_key']}）可用 ID 见 references/field-map.md。"
            + (f"\n（已删除上一轮遗留的 {args.out}）" if os.path.exists(args.out) else "")
        )

    items = []
    for f in spec.get("fields", []):
        items.append(("field", f))
    for c in spec.get("cells", []):
        items.append(("cell", c))

    ops, rows, errors, warns = [], [], [], []

    # 先摸一遍文档现状：已填项决定了「必填未提供」到底是错误还是无需再填
    cur = {}
    for kind, f in items:
        cur[f["id"]] = (analyse_field(doc, f, units) if kind == "field"
                        else analyse_cell(doc, f, units))

    # 判断这份文档是不是「还是模板」：模板里绝大多数待填项都应是空白。
    # 用来区分「正则匹配不到 = 已经填过」和「正则匹配不到 = 模板改版了」。
    blanks = sum(1 for _k, f in items if cur[f["id"]]["status"] == "BLANK")
    pristine = blanks >= max(2, len(items) // 3)

    for kind, f in items:
        fid = f["id"]
        explicit = fid in values
        val = values.get(fid)
        if val is None and f.get("default") is not None:
            val = f["default"]
        if isinstance(val, str):
            val = val.strip()
            if f.get("vtype") in NUMERIC_VTYPES:
                # 全角数字/全角百分号统一成半角，否则「２０」会被判为非法或被原样写进合同
                val = unicodedata.normalize("NFKC", val)
        if val is None:
            if cur[fid]["status"] == "FILLED":
                rows.append((fid, f["label"], "—", "— 文档中已有值，无需再填"))
            elif f.get("required"):
                errors.append(f"{fid} {f['label']}：必填项未提供")
                rows.append((fid, f["label"], "—", "✗ 必填未提供"))
            else:
                rows.append((fid, f["label"], "—", "— 未提供，留空"))
            continue
        ok, msg = validate_one(f, val)
        if not ok:
            errors.append(f"{fid} {f['label']}：{msg}")
            rows.append((fid, f["label"], val, f"✗ {msg}"))
            continue

        if kind == "field":
            rec = cur[fid]
            # 模板态下正则匹配不到，说明模板这一句改过了——不能当成"已填好"放过去
            if rec["status"] == "FILLED" and pristine and f.get("required"):
                errors.append(
                    f"{fid} {f['label']}：模板里该处已不再是空白，疑似模板改版或填错了文件；"
                    f"请用 标准版原始模板重新生成"
                )
                rows.append((fid, f["label"], val, "✗ 模板疑似改版"))
                continue
            # DEFAULT：模板预置值。用户没显式给值就保持原样，给了值才覆盖。
            if rec["status"] == "FILLED" or (rec["status"] == "DEFAULT" and not explicit):
                rendered = render(f["render"], val, f.get("vtype"))
                if rec["status"] == "FILLED" and explicit and rendered not in (rec.get("current") or ""):
                    # 文档里已有值，且和新给的值不一样——不能静默吞掉，必须让用户选
                    errors.append(
                        f"{fid} {f['label']}：文档该处已有内容「{str(rec.get('current'))[:24]}」，"
                        f"与本次提供的「{val}」不一致。请确认是保留原值还是覆盖"
                    )
                    rows.append((fid, f["label"], val, "✗ 与文档现值冲突"))
                    continue
                warns.append(f"{fid} {f['label']}：文档中已有值（{rec['status']}），保持原样")
                # 用户没显式给值、用的是 spec 里登记的模板预置值时，
                # 对照表要写「按默认」，避免看起来像「把 10000 写进去了」。
                shown = val if explicit else "按默认"
                rows.append((fid, f["label"], shown, "— 保持现值"))
                continue
            if rec["status"] != "BLANK" and rec["status"] != "DEFAULT":
                msg = f"{fid} {f['label']}：无法在模板中定位（{rec['status']}）"
                if f.get("required"):
                    errors.append(msg + "，必填项不能跳过，请核对模板版本")
                    rows.append((fid, f["label"], val, "✗ 定位失败"))
                else:
                    warns.append(msg + "，已跳过")
                    rows.append((fid, f["label"], val, "— 跳过（定位失败）"))
                continue
            find = rec["raw"]
            new = render(f["render"], val, f.get("vtype"))
            mode = "insert" if f.get("mode") == "append" else "replace"
            ops.append({"field_id": fid, "kind": mode, "find": find,
                        "occurrence": int(f["target"].get("occurrence", 1)),
                        "value": new, "expect": new, "ctx": rec.get("context", "")})
            rows.append((fid, f["label"], val, new))
        else:
            rec = cur[fid]
            if rec["status"] == "FILLED":
                warns.append(f"{fid} {f['label']}：该格已有内容，保持原样")
                rows.append((fid, f["label"], val, "— 保持现值"))
                continue
            if rec["status"] != "BLANK":
                msg = f"{fid} {f['label']}：无法在模板该单元格定位（{rec['status']}）"
                if f.get("required"):
                    errors.append(msg + "，必填项不能跳过")
                    rows.append((fid, f["label"], val, "✗ 定位失败"))
                else:
                    warns.append(msg + "，已跳过")
                    rows.append((fid, f["label"], val, "— 跳过（定位失败）"))
                continue
            text = rec["para_text"]
            frag = render(f["render"], val, f.get("vtype"))
            if f.get("mode") == "append":
                new_cell = text + frag
            else:
                m = re.search(f["pattern"], text)
                new_cell = text[:m.start()] + frag + text[m.end():]
            ops.append({"field_id": fid, "kind": "cell", "table": f["table"],
                        "row": f["row"], "col": f["col"], "value": new_cell,
                        "expect": new_cell})
            rows.append((fid, f["label"], val, new_cell))

    # 网格表
    for g in spec.get("grids", []):
        gid = g["id"]
        gval = values.get(gid)
        if not gval:
            if g.get("required"):
                rec = analyse_grid(doc, g)
                if rec.get("filled_cells", 0) > 0:
                    rows.append((gid, g["label"], "—", "— 表中已有数据，无需再填"))
                else:
                    errors.append(
                        f"{gid} {g['label']}：必填项未提供（至少给出季度或月度中的一档）")
                    rows.append((gid, g["label"], "—", "✗ 必填未提供"))
            continue
        plans = [("quarter", g["rows"]["quarter"], g["quarter_cols"], gval.get("quarter") or []),
                 ("month", g["rows"]["month"], g["month_cols"], gval.get("month") or [])]
        # 数组长度必须精确：多给的会被静默丢弃，少给的会被当成空白
        for key, row, cols, arr in plans:
            if arr and len(arr) != len(cols):
                errors.append(
                    f"{gid} {g['label']}：{key} 应给 {len(cols)} 个数，实际给了 {len(arr)} 个"
                )
        for key, row, cols, arr in plans:
            for i, ci in enumerate(cols):
                if i >= len(arr) or str(arr[i]).strip() == "":
                    continue
                v = str(arr[i]).strip()
                if not re.fullmatch(r"\d+(\.\d+)?", v):
                    errors.append(f"{gid} {g['label']} {key}#{i + 1}：应为数字，当前「{v}」")
                    continue
                ops.append({"field_id": f"{gid}-{key}{i + 1}", "kind": "cell",
                            "table": g["table"], "row": row, "col": ci,
                            "value": v, "expect": v})
        filled = sum(1 for k, _, _, arr in plans for x in arr if str(x).strip())
        rows.append((gid, g["label"], f"{filled} 格", "—"))

    # ------------------------------------------------------------ 跨字段一致性
    # 严格按当前规格（JX / DX）取键，不能跨文件回退：JX 与 DX 的段落编号不同，
    # 用 `or` 回退会让 DX 的校验读到 JX 的值（或反之），等于拿 A 合同的日期校验 B 合同。
    DOC_KEYS = {
        "JX": {"no": "JX-01", "start": "JX-09", "end": "JX-10", "sign": "JX-20",
               "num": "JX-12", "cn": "JX-11", "q": "JX-18", "y": "JX-19",
               "prov": "JX-06", "city": "JX-07"},
        "DX": {"no": "DX-01", "start": "DX-09", "end": "DX-10", "sign": "DX-23",
               "num": "DX-12", "cn": "DX-11", "q": "DX-21", "y": "DX-22",
               "prov": "DX-06", "city": "DX-07"},
    }
    K = DOC_KEYS.get(spec["doc_key"], {})

    def gv(k):
        fid = K.get(k)
        if not fid or values.get(fid) is None:
            return ""
        return str(values.get(fid)).strip()

    # 主合同编号：类型码必须与合同类型一致，年份必须等于签订年份
    ce, cw = check_contract_no_consistency(gv("no"), spec["doc_key"], gv("sign"))
    errors.extend(ce)
    warns.extend(cw)

    d9, d10 = gv("start"), gv("end")
    if d9 and d10 and check_date(d9)[0] and check_date(d10)[0] and d9 >= d10:
        errors.append(f"合同期限：起始日 {d9} 必须早于终止日 {d10}")
    sign = gv("sign")
    if sign and d9 and check_date(sign)[0] and check_date(d9)[0] and sign > d9:
        warns.append(f"签订时间 {sign} 晚于合同起始日 {d9}（倒签），请确认是否有意为之")
    if sign and d10 and check_date(sign)[0] and check_date(d10)[0] and sign > d10:
        errors.append(f"签订时间 {sign} 晚于合同终止日 {d10}，请核对")

    # 保证金：大小写必须成对给，只给一半会静默漏填另一半
    num, cn = gv("num"), gv("cn")
    if bool(num) != bool(cn):
        errors.append(
            f"履约保证金：大写与小写必须同时给（当前只有"
            f"{'小写 ' + num if num else '大写 ' + cn}），否则另一半会保持模板默认值"
        )
    elif num and cn:
        ok, msg = check_amount_cn(cn, int(num) if num.isdigit() else None)
        if not ok:
            errors.append(f"履约保证金：{msg}（小写 {num} 元）")

    q, y = gv("q"), gv("y")
    try:
        if q and y and float(y) > float(q) * 2:
            warns.append(f"年度返利 {y}% 明显高于季度返利 {q}%，请确认是否填反")
    except ValueError:
        pass

    prov, city = gv("prov"), gv("city")
    if prov in ("北京", "上海", "天津", "重庆"):
        warns.append(
            f"「{prov}」为直辖市：模板句式是「＿＿省＿＿市」，填「{prov}省{city}市」读感不佳，"
            f"建议与甲方确认写法（如「{prov}市{city}区」或只在市级填「{prov}」）"
        )

    g1 = values.get(f"{spec['doc_key']}-G1")
    if g1:
        for key, label in (("quarter", "季度"), ("month", "月度")):
            arr = [str(x).strip() for x in (g1.get(key) or []) if str(x).strip()]
            if arr:
                try:
                    tot = sum(float(x) for x in arr)
                    warns.append(f"销售目标{label}分解合计 {tot:g} 万元（仅供参考，未做自动平衡）")
                except ValueError:
                    pass

    digest = hashlib.sha256(
        json.dumps(values, ensure_ascii=False, sort_keys=True).encode("utf-8")
    ).hexdigest()
    plan = {"doc": spec["doc_key"], "spec": os.path.abspath(args.spec),
            "template": os.path.abspath(args.doc),
            "spec_sha256": file_sha256(args.spec),
            "base_sha256": file_sha256(args.doc),
            "values_digest": digest,
            "source_values": values,
            "ops": ops}

    print(f"合同：{spec.get('title')}")
    print(f"模板：{os.path.basename(args.doc)}")
    print("-" * 78)
    for fid, label, val, new in rows:
        print(f"{fid:<8}{label:<22}入参={str(val)[:22]:<24}落笔={str(new)[:34]}")
    if warns:
        print("-" * 78)
        print("提示：")
        for w in warns:
            print("  ! " + w)
    if errors:
        print("-" * 78)
        print("必须修正：")
        for e in errors:
            print("  ✗ " + e)
        print("-" * 78)
        # fail-closed：存在错误时不产出可执行计划，并清掉上一轮的旧计划，避免被误执行
        rejected = args.out + ".rejected"
        json.dump(plan, open(rejected, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        if os.path.exists(args.out):
            os.remove(args.out)
            print(f"已删除上一轮遗留的 {args.out}")
        print(f"存在 {len(errors)} 处错误，未生成可执行计划（仅供排查：{rejected}）。")
        return 1
    # 本轮通过后，清掉上一轮失败留下的排查文件，避免被误当成可执行计划使用。
    if os.path.exists(args.out + ".rejected"):
        os.remove(args.out + ".rejected")
    json.dump(plan, open(args.out, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print("-" * 78)
    print(f"校验通过，已生成 {args.out}，共 {len(ops)} 条操作。")
    print(f"值指纹 {digest[:16]}…（验收时凭此确认计划未被改动）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
