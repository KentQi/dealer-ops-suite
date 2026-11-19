#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""合同包审阅（确定性部分）。

审阅对象：主合同 values + 防窜货附件 values（可选）+ 价格政策表（可选）。
输出：统一格式的「需确认项」JSON / Markdown。

它不是合同的法律意见，也不替代用户决策。它只做三件事：
  1. 找客观矛盾：必填缺失、日期倒挂、比例越界、目标合计不一致、附件日期不一致；
  2. 把关键条款的缺口变成可确认的问题；
  3. 给 LLM 语义审阅提供确定的输入与合并入口。

用法：
    python3 review_contract_package.py \
        --spec spec/jx.json --values /tmp/values.json \
        --fc-values /tmp/fc-values.json \
        --policy "价格政策.xlsx" \
        --out /tmp/合同审阅.json --md /tmp/合同审阅.md

    # 合并 LLM 语义审阅结果（schema 见 references/合同包审阅.md）
    python3 review_contract_package.py ... --llm-findings /tmp/llm-findings.json

退出码：0 = 无阻断项；1 = 存在阻断项；2 = 输入错误。
"""
import argparse
import json
import os
import re
import subprocess
import sys
import unicodedata
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
SEVERITY_ORDER = {"blocker": 0, "high": 1, "medium": 2, "low": 3}
SEVERITY_LABEL = {"blocker": "阻断", "high": "高", "medium": "中", "low": "低"}


def fail(message, code=2):
    print(f"[中止] {message}", file=sys.stderr)
    sys.exit(code)


def load_json(path):
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception as exc:
        fail(f"读取 JSON 失败：{path}\n{exc}")


def blank(value):
    return value is None or (isinstance(value, str) and not value.strip())


def text(value):
    return "" if value is None else str(value).strip()


def norm(value):
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", text(value))).lower()


def as_float(value):
    try:
        return float(str(value).strip().rstrip("%"))
    except (TypeError, ValueError):
        return None


def valid_date(value):
    try:
        datetime.strptime(text(value), "%Y-%m-%d")
        return True
    except ValueError:
        return False


def valid_uscc(value):
    chars = "0123456789ABCDEFGHJKLMNPQRTUWXY"
    weights = [1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28]
    code = text(value).upper()
    if len(code) != 18 or any(c not in chars for c in code):
        return False
    total = sum(chars.index(code[i]) * weights[i] for i in range(17))
    check = 31 - total % 31
    check = 0 if check == 31 else check
    return chars[check] == code[17]


def valid_phone(value):
    v = text(value)
    return bool(re.fullmatch(r"1[3-9]\d{9}", v) or re.fullmatch(r"0\d{2,3}-?\d{7,8}(-\d+)?", v))


def valid_person(value):
    v = text(value)
    return 2 <= len(v) <= 15 and not re.search(r"[\d()（）]", v)


def valid_company(value):
    v = text(value)
    return len(v) >= 4 and bool(re.search(r"(公司|企业|商行|经营部|中心|店|厂|社|事务所|合伙)$", v))


def valid_address(value):
    v = text(value)
    return len(v) >= 8 and bool(re.search(r"(省|市|区|县|旗|镇)", v)) and bool(re.search(r"(路|街|号|楼|室|层|栋|园区|仓)", v))


def valid_contract_no(value):
    v = text(value)
    if v == "待定":
        return True
    return bool(re.fullmatch(r"[A-Z]{2}-[A-Z]{2}(?:-[A-Z]{2})?-\d{4}-\d{3}", v))


def check_value(vtype, value):
    """返回错误文本；None 表示通过。只做审阅级校验，精确校验仍以 build_plan.py 为准。"""
    v = text(value)
    if vtype == "contract_no":
        return None if valid_contract_no(v) else "编号格式应为「品牌码-类型码-年份-序号」，或填「待定」"
    if vtype == "uscc":
        return None if valid_uscc(v) else "统一社会信用代码应为 18 位且校验位正确"
    if vtype == "date":
        return None if valid_date(v) else "日期应为 YYYY-MM-DD"
    if vtype == "phone":
        return None if valid_phone(v) else "电话格式不正确"
    if vtype == "person":
        return None if valid_person(v) else "姓名应为 2-15 字，不含数字与括号"
    if vtype == "company":
        return None if valid_company(v) else "公司全称应以有限公司/合伙企业/商行等主体类型结尾"
    if vtype == "address":
        return None if valid_address(v) else "地址应包含省市区与详细门牌"
    if vtype == "region":
        return None if 1 <= len(v) <= 10 else "行政区划名称过长"
    if vtype == "percent_rebate":
        n = as_float(v)
        return None if n is not None and 0 <= n <= 30 else "返利比例应在 0-30 之间"
    if vtype == "percent_any":
        n = as_float(v)
        return None if n is not None and 0 < n <= 100 else "比例应在 0-100 之间"
    if vtype == "amount_num":
        return None if re.fullmatch(r"\d{1,9}", v) else "金额应为不超过 9 位的整数（元）"
    if vtype == "amount_wan":
        n = as_float(v)
        return None if n is not None and n > 0 else "金额应大于 0（万元）"
    if vtype == "day":
        return None if v.isdigit() and 1 <= int(v) <= 28 else "日期应为 1-28 的整数"
    if vtype == "int":
        return None if v.isdigit() else "应为非负整数"
    return None


def add(findings, category, severity, message, field="", label="", value="", ask=""):
    findings.append({
        "id": "",
        "category": category,
        "severity": severity,
        "field": field,
        "label": label,
        "message": message,
        "evidence": {"value": value},
        "ask": ask or "请确认这一项是否按当前内容执行。",
    })


def check_fields(spec, values, findings):
    for f in spec.get("fields", []):
        fid = f.get("id", "")
        label = f.get("label", fid)
        value = values.get(fid)
        if blank(value):
            if f.get("required"):
                add(findings, "必填/完整性", "blocker", "必填项未提供", fid, label, "",
                    "请补齐后重跑；若模板已有默认值，请说明并确认。")
            continue
        error = check_value(f.get("vtype", "text"), value)
        if error:
            add(findings, "格式/取值", "blocker", error, fid, label, text(value),
                "请修正该字段后重跑。")
        if f.get("vtype") == "contract_no" and text(value) == "待定":
            add(findings, "关键条款提醒", "medium", "合同编号仍是「待定」", fid, label, text(value),
                "签署前必须替换为正式编号；否则无法作为合同包索引。")

    for c in spec.get("cells", []):
        cid = c.get("id", "")
        label = c.get("label", cid)
        value = values.get(cid)
        if blank(value):
            if c.get("required"):
                add(findings, "必填/完整性", "blocker", "必填项未提供", cid, label, "",
                    "请补齐后重跑；若模板已有默认值，请说明并确认。")
            continue
        error = check_value(c.get("vtype", "int"), value)
        if error:
            add(findings, "格式/取值", "blocker", error, cid, label, text(value),
                "请修正该字段后重跑。")


def check_grids(spec, values, findings):
    for g in spec.get("grids", []):
        gid = g.get("id", "")
        label = g.get("label", gid)
        gval = values.get(gid) or {}
        if not isinstance(gval, dict):
            add(findings, "目标/指标", "high", "目标表值不是对象（应含 quarter/month 数组）", gid, label, text(gval),
                "请按 spec 的数组结构提供季度或月度目标。")
            continue
        quarter = gval.get("quarter") or []
        month = gval.get("month") or []
        if not quarter and not month:
            add(findings, "目标/指标", "high", "未提供目标表数据", gid, label, "",
                "若模板已有目标值，请说明；否则至少给出季度或月度中的一档。")
            continue
        if quarter and len(quarter) != 4:
            add(findings, "目标/指标", "blocker", f"季度目标应给 4 个数，实际 {len(quarter)} 个", gid, label, str(quarter),
                "请补齐四个季度目标。")
        if month and len(month) != 12:
            add(findings, "目标/指标", "blocker", f"月度目标应给 12 个数，实际 {len(month)} 个", gid, label, str(month),
                "请补齐十二个月目标。")
        if quarter and month:
            qsum, msum = sum(as_float(x) or 0 for x in quarter), sum(as_float(x) or 0 for x in month)
            if abs(qsum - msum) > 0.01:
                add(findings, "目标/指标", "high",
                    f"季度目标合计 {qsum:g} 与月度目标合计 {msum:g} 不一致", gid, label,
                    f"quarter={quarter}; month={month}",
                    "请确认以哪一档为准，或修正另一档。")


def check_dates(spec, values, findings):
    key = spec.get("doc_key", "")
    prefix = "JX" if key == "JX" else "DX" if key == "DX" else ""
    if not prefix:
        return
    start = text(values.get(f"{prefix}-09"))
    end = text(values.get(f"{prefix}-10"))
    sign = text(values.get(f"{prefix}-20" if prefix == "JX" else f"{prefix}-23"))
    if start and end and valid_date(start) and valid_date(end) and start >= end:
        add(findings, "日期逻辑", "blocker", "合同起始日不早于终止日", f"{prefix}-09/{prefix}-10",
            "合同期限", f"{start} → {end}", "请确认合同期限是否正确。")
    if start and sign and valid_date(start) and valid_date(sign) and sign > start:
        add(findings, "日期逻辑", "high", "签订时间晚于合同起始日", f"{prefix}-20" if prefix == "JX" else f"{prefix}-23",
            "签订时间", f"签订 {sign} / 起始 {start}", "请确认合同是否倒签，或调整起始日。")


def check_guarantee(spec, values, findings):
    prefix = "JX" if spec.get("doc_key") == "JX" else "DX" if spec.get("doc_key") == "DX" else ""
    if not prefix:
        return
    cn = values.get(f"{prefix}-11")
    num = values.get(f"{prefix}-12")
    if blank(cn) and blank(num):
        add(findings, "保证金", "medium", "未设置履约保证金", f"{prefix}-11/{prefix}-12",
            "履约保证金", "", "请确认业务上是否允许不设保证金；若允许，请留痕。")
    elif blank(cn) or blank(num):
        add(findings, "保证金", "high", "保证金大写/小写只给了一半", f"{prefix}-11/{prefix}-12",
            "履约保证金", f"大写={text(cn)} 小写={text(num)}", "请同时补齐大写与小写，避免只落笔一半。")


def check_rebate(spec, values, findings):
    prefix = "JX" if spec.get("doc_key") == "JX" else "DX" if spec.get("doc_key") == "DX" else ""
    if not prefix:
        return
    q = as_float(values.get(f"{prefix}-18" if prefix == "JX" else f"{prefix}-21"))
    y = as_float(values.get(f"{prefix}-19" if prefix == "JX" else f"{prefix}-22"))
    if q is not None and y is not None:
        if y > q * 2:
            add(findings, "返利", "high", "年度返利明显高于季度返利", f"{prefix}-18/{prefix}-19",
                "返利", f"季度 {q:g}% / 年度 {y:g}%", "请确认是否填反，或说明两档口径不同。")
        if q == 0 and y == 0:
            add(findings, "返利", "medium", "季度与年度返利均为 0", f"{prefix}-18/{prefix}-19",
                "返利", "0 / 0", "请确认这是政策选择，还是漏填。")


def check_fc(fc_spec, fc_values, main_spec, main_values, findings):
    if not fc_spec or not fc_values:
        add(findings, "附件/跨文件", "medium", "未提供防窜货附件审阅", "FC", "附件一",
            "", "请把附件签署日期纳入合同包审阅，确认与主合同签订时间一致。")
        return
    main_sign = text(main_values.get("JX-20") or main_values.get("DX-23"))
    for cid, label in (("FC-C1", "附件签署日期-甲方"), ("FC-C2", "附件签署日期-乙方")):
        value = text(fc_values.get(cid))
        if not value:
            add(findings, "附件/跨文件", "blocker", "附件签署日期未提供", cid, label, "",
                "附件必须显式签署日期，不能默认跟随主合同。")
        elif not valid_date(value):
            add(findings, "附件/跨文件", "blocker", "附件签署日期格式不正确", cid, label, value,
                "请改为 YYYY-MM-DD。")
        elif main_sign and valid_date(main_sign) and value != main_sign:
            add(findings, "附件/跨文件", "high", "附件签署日期与主合同签订时间不一致", cid, label,
                f"附件 {value} / 主合同 {main_sign}", "请确认以哪个日期为准，并同步修改。")


def check_policy(policy_path, findings):
    if not policy_path:
        add(findings, "政策表", "medium", "未提供价格政策表", "policy", "价格政策",
            "", "本次合同包审阅未覆盖价格/出货政策；请补跑政策表自检。")
        return
    script = os.path.join(HERE, "check_policy.py")
    if not os.path.isfile(script):
        add(findings, "政策表", "high", "找不到 check_policy.py，无法审阅政策表", "policy", "价格政策",
            policy_path, "请确认 Skill 安装完整。")
        return
    proc = subprocess.run([sys.executable, script, policy_path, "--json"],
                          capture_output=True, text=True)
    if proc.returncode not in (0, 2):
        add(findings, "政策表", "high", "政策表无法解析", "policy", "价格政策", policy_path,
            (proc.stderr or proc.stdout or "check_policy.py 执行失败").strip()[:300])
        return
    try:
        data = json.loads(proc.stdout)
    except Exception:
        add(findings, "政策表", "high", "政策表检查结果不是合法 JSON", "policy", "价格政策", policy_path,
            (proc.stdout or "").strip()[:300])
        return
    for item in data.get("needs_confirmation", []):
        label = item.get("type") or item.get("id") or "政策表需确认项"
        add(findings, "政策表", "high", f"[{item.get('id', '')}] {label}", "policy", "价格政策",
            json.dumps(item, ensure_ascii=False), item.get("ask", "请确认政策表以哪一行为准。"))


def merge_llm(path, findings):
    if not path:
        return {"required": True, "status": "not_run", "note": "仅完成确定性审阅；LLM 语义审阅未运行。"}
    data = load_json(path)
    items = data.get("findings", data if isinstance(data, list) else [])
    if not isinstance(items, list):
        fail("LLM findings JSON 必须是数组，或包含 findings 数组")
    for item in items:
        if not isinstance(item, dict):
            continue
        severity = item.get("severity", "high")
        if severity not in SEVERITY_ORDER:
            severity = "high"
        findings.append({
            "id": "",
            "category": item.get("category", "LLM 语义审阅"),
            "severity": severity,
            "field": item.get("field", ""),
            "label": item.get("label", ""),
            "message": item.get("message") or item.get("issue") or "LLM 提出需确认项",
            "evidence": item.get("evidence", {}),
            "ask": item.get("ask") or item.get("question") or "请确认这一条款如何执行。",
        })
    return {"required": True, "status": "merged", "note": f"合并 LLM 语义审阅 {len(items)} 项。"}


def summarize(findings):
    summary = {"blocker": 0, "high": 0, "medium": 0, "low": 0, "total": len(findings)}
    for item in findings:
        summary[item["severity"]] += 1
    return summary


def render_markdown(result):
    lines = [
        "# 合同包审阅",
        "",
        f"- 规格：`{result['inputs'].get('spec', '')}`",
        f"- 值文件：`{result['inputs'].get('values', '')}`",
        f"- 附件值：`{result['inputs'].get('fc_values') or '未提供'}`",
        f"- 政策表：`{result['inputs'].get('policy') or '未提供'}`",
        f"- 生成时间：{result['generated_at']}",
        "",
        f"## 汇总：阻断 {result['summary']['blocker']}　高 {result['summary']['high']}　"
        f"中 {result['summary']['medium']}　低 {result['summary']['low']}　合计 {result['summary']['total']}",
        "",
    ]
    for severity in ("blocker", "high", "medium", "low"):
        items = [x for x in result["findings"] if x["severity"] == severity]
        if not items:
            continue
        lines += [f"## {SEVERITY_LABEL[severity]}（{len(items)}）", ""]
        for item in items:
            where = f" `{item['field']}`" if item.get("field") else ""
            value = item.get("evidence", {}).get("value")
            value_text = f"　当前值：`{value}`" if value not in (None, "") else ""
            lines.append(f"- **{item['category']}**{where}：{item['message']}{value_text}")
            lines.append(f"  - 需确认：{item['ask']}")
        lines.append("")
    lines += ["## LLM 语义审阅", "", result["llm_review"]["note"], ""]
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="合同包审阅（确定性）")
    ap.add_argument("--spec", required=True, help="主合同 spec，如 spec/jx.json")
    ap.add_argument("--values", required=True, help="主合同 values.json")
    ap.add_argument("--fc-spec", help="附件一 spec，如 spec/fc.json")
    ap.add_argument("--fc-values", help="附件一 values.json")
    ap.add_argument("--policy", help="价格政策表 .xlsx（可选）")
    ap.add_argument("--llm-findings", help="LLM 语义审阅 findings JSON（可选）")
    ap.add_argument("--out", help="输出 JSON 路径")
    ap.add_argument("--md", help="输出 Markdown 路径")
    ap.add_argument("--json", action="store_true", help="stdout 输出 JSON")
    ap.add_argument("--fail-on-high", action="store_true", help="存在 high 时也以 1 退出")
    args = ap.parse_args()

    spec = load_json(args.spec)
    values = load_json(args.values)
    if not isinstance(values, dict):
        fail("values.json 顶层必须是对象")
    fc_spec = load_json(args.fc_spec) if args.fc_spec else None
    fc_values = load_json(args.fc_values) if args.fc_values else None
    if fc_values is not None and not isinstance(fc_values, dict):
        fail("附件 values.json 顶层必须是对象")

    findings = []
    check_fields(spec, values, findings)
    check_grids(spec, values, findings)
    check_dates(spec, values, findings)
    check_guarantee(spec, values, findings)
    check_rebate(spec, values, findings)
    check_fc(fc_spec, fc_values, spec, values, findings)
    check_policy(args.policy, findings)
    llm_review = merge_llm(args.llm_findings, findings)

    findings.sort(key=lambda x: (SEVERITY_ORDER[x["severity"]], x["category"], x["field"]))
    for index, item in enumerate(findings, 1):
        item["id"] = f"C-{index:03d}"

    result = {
        "schema": "dealer-ops/contract-review@1",
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "inputs": {
            "spec": args.spec,
            "values": args.values,
            "fc_spec": args.fc_spec,
            "fc_values": args.fc_values,
            "policy": args.policy,
            "llm_findings": args.llm_findings,
        },
        "summary": summarize(findings),
        "findings": findings,
        "llm_review": llm_review,
    }

    markdown = render_markdown(result)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            json.dump(result, fh, ensure_ascii=False, indent=2)
            fh.write("\n")
    if args.md:
        with open(args.md, "w", encoding="utf-8") as fh:
            fh.write(markdown)
            fh.write("\n")

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(markdown)
        if args.out:
            print(f"\nJSON：{args.out}")
        if args.md:
            print(f"Markdown：{args.md}")

    summary = result["summary"]
    if summary["blocker"] or (args.fail_on_high and summary["high"]):
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
