"""按当前期间与有效合同生成对账引擎 RECON_CFG。

铺货代销读取 BIP 期初、乙方销售和库存；月结买断读取签收、退货、动销。
指定期间目录不存在即中止，不借用其他月份；多份或零份有效合同均中止。
stdout 仅输出单行 JSON，告警和错误写 stderr。
"""
import argparse
import datetime
import json
import os
import re
import sys

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import koujing
from settlement_rules import RuleViolation, normalize_business_type, parse_day


def read_sheet(wb, name):
    if name not in wb.sheetnames:
        return []
    ws = wb[name]
    headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]
    return [dict(zip(headers, row)) for row in ws.iter_rows(min_row=2, values_only=True)
            if any(value is not None and str(value).strip() for value in row)]


def month_bounds(period):
    year, month = int(period[:4]), int(period[4:])
    if not 2000 <= year <= 2100 or not 1 <= month <= 12:
        raise ValueError(f"期间 {period!r} 非法：须为合法 YYYYMM（如 202611）")
    start = datetime.date(year, month, 1)
    next_month = datetime.date(year + (month == 12), 1 if month == 12 else month + 1, 1)
    return start, next_month - datetime.timedelta(days=1), next_month


def _date_text(value):
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.strftime("%Y-%m-%d")
    return str(value).strip() if value is not None else ""


def _as_int(value):
    try:
        number = float(str(value).strip())
    except (TypeError, ValueError):
        return None
    return int(number) if number == int(number) else None


def _pay_day_from_text(value):
    """仅提取明确日号；解析不到绝不默认 5 日。"""
    match = re.search(r"([0-9]{1,2})\s*日", str(value or ""))
    if match and 1 <= int(match.group(1)) <= 28:
        return int(match.group(1))
    return None


def _excel_files(directory):
    try:
        return sorted(name for name in os.listdir(directory)
                      if name.lower().endswith((".xls", ".xlsx", ".xlsm")) and not name.startswith("~$"))
    except OSError:
        return []


def _fail(message):
    raise SystemExit(message)


def select_contract(contracts, start, end):
    """只选完整覆盖本期的唯一合同；没有日期就不算命中。"""
    matches = []
    for index, contract in enumerate(contracts, 2):
        try:
            contract_start = parse_day(contract.get("起始"), "合同起始日")
            contract_end = parse_day(contract.get("终止"), "合同终止日")
        except RuleViolation:
            continue
        if contract_start <= start and end <= contract_end:
            matches.append((index, contract))
    if not matches:
        _fail(f"[中止] 没有完整覆盖对账期间 {start}—{end} 的有效合同；不默认使用第一份合同。")
    if len(matches) > 1:
        _fail("[中止] 多份合同同时覆盖本对账期间，行号："
              + "、".join(str(index) for index, _ in matches) + "；请先明确唯一有效合同。")
    return matches[0][1]


def _load_rows(path, label):
    if not path:
        return []
    try:
        with open(path, encoding="utf-8") as handle:
            rows = json.load(handle)
    except (OSError, ValueError) as exc:
        _fail(f"[{label}读取失败] {exc}")
    if not isinstance(rows, list) or not all(isinstance(row, dict) for row in rows):
        _fail(f"[{label}格式错误] 须为流水对象数组")
    return rows


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("master")
    parser.add_argument("--period", required=True, help="YYYYMM")
    parser.add_argument("--asof", help="库存／挂账导出日，默认次月 1 日")
    parser.add_argument("--pay", help="铺货代销付款日 YYYY-MM-DD；月结买断必须在主档登记付款截止日，不接受此参数")
    parser.add_argument("--base", help="本期目录；缺省固定为主档同级的 YYYYMM")
    parser.add_argument("--out", help="输出工作簿全路径")
    parser.add_argument("--statement-no", help="可选：出账前已登记的对账单编号；不自动生成或沿用其他期间编号")
    parser.add_argument("--strict", action="store_true")
    parser.add_argument("--scene", default="每月出账", choices=list(koujing.SCENES))
    parser.add_argument("--koujing", default="")
    parser.add_argument("--allow-legacy", action="store_true", help="仅历史铺货只读复核可绕过口径问卡")
    parser.add_argument("--mode", default="full", choices=("full", "incremental"))
    parser.add_argument("--in-workbook", default="")
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--pdf", action="store_true")
    parser.add_argument("--pdf-sheets", default="00,01,02,08")
    parser.add_argument("--original-sales", help="历史销售流水 JSON，用于退货追溯")
    parser.add_argument("--original-receipts", help="历史签收流水 JSON，用于月结退货追溯")
    parser.add_argument("--cfg-out", help="将同一份 RECON_CFG 原子写入 JSON 文件，供 apply_v17.py --cfg 使用")
    args = parser.parse_args()

    if not re.fullmatch(r"[0-9]{6}", args.period or "") or not 1 <= int(args.period[4:]) <= 12 \
            or not 2000 <= int(args.period[:4]) <= 2100:
        _fail("[参数错误] --period 须为合法 YYYYMM（如 202611）")
    try:
        start, end, next_month = month_bounds(args.period)
    except ValueError as exc:
        _fail(f"[参数错误] {exc}")
    asof = args.asof or next_month.isoformat()
    try:
        asof_day = datetime.date.fromisoformat(asof)
        if args.pay:
            datetime.date.fromisoformat(args.pay)
    except ValueError:
        _fail("[参数错误] --asof / --pay 须为 YYYY-MM-DD")
    try:
        wb = openpyxl.load_workbook(args.master, data_only=True)
    except Exception as exc:
        _fail(f"[无法读取主档] {args.master}：{exc}")
    try:
        parties, contracts = read_sheet(wb, "主体"), read_sheet(wb, "合同")
    finally:
        wb.close()
    party_a = next((row for row in parties if str(row.get("角色", "")).strip() == "甲方"), {})
    party_b = next((row for row in parties if str(row.get("角色", "")).strip() == "乙方"), {})
    contract = select_contract(contracts, start, end)
    try:
        business_type = normalize_business_type(contract.get("业务类型") or contract.get("结算方式"))
    except RuleViolation as exc:
        _fail(f"[业务路由阻断] {exc.message}")
    if business_type == "monthly_buyout" and args.pay:
        _fail("[参数错误] 月结买断付款截止日必须在主档合同登记，不能用 --pay 补录")

    payment_due = _date_text(contract.get("付款截止日"))
    if business_type == "monthly_buyout" and not payment_due:
        # 仅兼容已经是完整日期的旧字段，不解析或虚构“次月某日”。
        try:
            payment_due = parse_day(contract.get("付款日"), "付款日").isoformat()
        except RuleViolation:
            payment_due = ""
    payment_due = args.pay or payment_due
    contract_cfg = {
        "no": str(contract.get("合同编号") or "").strip(),
        "start": _date_text(contract.get("起始")), "end": _date_text(contract.get("终止")),
        "method": str(contract.get("业务类型") or contract.get("结算方式") or "").strip(),
        "basis": str(contract.get("结算基数") or "").strip(),
        "pay_day_text": str(contract.get("付款日") or "").strip(),
        "objection_days": _as_int(contract.get("异议期(工作日)")),
        "signback_days": _as_int(contract.get("签回期(工作日)")),
        "price_annex": str(contract.get("价格附件") or "").strip(),
        "buyout_date": _date_text(contract.get("买断日")),
        "payment_due_date": payment_due,
    }
    problems = []
    if not party_a.get("名称"):
        problems.append("主体表缺少甲方名称")
    if not party_b.get("名称"):
        problems.append("主体表缺少乙方名称")
    if not contract_cfg["no"]:
        problems.append("有效合同缺合同编号")
    if not party_a.get("开户行") or not party_a.get("账号"):
        problems.append("甲方开户行或账号为空")
    pay_day = _pay_day_from_text(contract_cfg["pay_day_text"])
    pay = args.pay or (next_month.replace(day=pay_day).isoformat() if pay_day else None)
    if business_type == "consignment" and not pay:
        problems.append("合同付款日无法识别；不会默认次月 5 日，可用 --pay 明确指定")
    if business_type == "monthly_buyout":
        for label, value in (("买断日", contract_cfg["buyout_date"]), ("付款截止日", payment_due)):
            try:
                parse_day(value, label)
            except RuleViolation:
                problems.append(f"月结买断必须明确{label}，不得默认或与另一日期合并")
    for problem in problems:
        print("[主档告警] " + problem, file=sys.stderr)
    if problems and (args.strict or business_type == "monthly_buyout"):
        _fail("[中止] 主档关键信息不全；请先补齐合同／主体关键字段。")

    root = os.path.dirname(os.path.abspath(args.master))
    base = os.path.abspath(args.base or os.path.join(root, args.period))
    if not os.path.isdir(base):
        _fail(f"[目录不存在] {base}；不会回退到最近月份或客户根目录。")

    def pick(rules, default):
        files = _excel_files(base)
        for include, exclude in rules:
            hits = [name for name in files if include in name and (exclude is None or exclude not in name)]
            if len(hits) > 1:
                _fail(f"[源文件歧义] 本期目录中“{include}”匹配多份文件："
                      + "、".join(sorted(hits)) + "；请只保留本期唯一有效源文件。")
            if hits:
                return hits[0]
        return default

    if business_type == "monthly_buyout":
        sources = {"receipts": pick([("签收", None)], "乙方-签收.xlsx"),
                   "returns": pick([("退货", None)], "乙方-退货.xlsx"),
                   "sales": pick([("动销", None), ("销售", "库存")], "乙方-动销.xlsx")}
        out_name = f"{args.period}月结买断对账.xlsx"
    else:
        sources = {"jia": pick([("挂账", None), ("BIP", None), ("期初", None)], "甲方-挂账.xls"),
                   "sales": pick([("销售数据", None), ("销售", "库存")], "乙方-销售数据.xlsx"),
                   "inventory": pick([("库存", None)], "乙方-库存.xlsx"),
                   "batch": pick([("效期", None), ("批次", None), ("生产日期", None)], ""),
                   "promo_use": pick([("使用", None), ("领用", None)], ""),
                   "barcode_map": pick([("条码对照", None), ("编码对照", None), ("物料编码", None)], ""),
                   "item_class": pick([("商品分类", None), ("分类", None)], "")}
        sources["expiry_ref"] = sources["batch"]
        out_name = f"{args.period}铺货代销对账.xlsx"

    mapping = None
    mapping_file = os.path.join(root, "_数据映射.json")
    if os.path.exists(mapping_file):
        try:
            with open(mapping_file, encoding="utf-8-sig") as handle:
                mapping = json.load(handle)
        except (OSError, ValueError) as exc:
            _fail(f"[映射文件损坏] {mapping_file}：{exc}")
        if not isinstance(mapping, dict) or mapping.get("schema") != "dealer-ops/data-mapping@1":
            _fail("[映射 schema 不支持] 要求 dealer-ops/data-mapping@1")

    cfg = {
        "business_type": business_type, "base_dir": base, "master_path": os.path.abspath(args.master),
        "statement_no": str(args.statement_no or "").strip(),
        "party_a": {"name": party_a.get("名称", ""), "tax": party_a.get("统一社会信用代码", ""),
                    "addr": party_a.get("地址", ""), "contact": party_a.get("联系人", ""),
                    "phone": party_a.get("电话", "")},
        "party_b": {"name": party_b.get("名称", ""), "tax": party_b.get("统一社会信用代码", ""),
                    "addr": party_b.get("地址", ""), "contact": party_b.get("联系人", ""),
                    "phone": party_b.get("电话", "")},
        "contract": contract_cfg,
        "bank": {"account_name": party_a.get("户名") or party_a.get("名称", ""),
                 "account_no": party_a.get("账号", ""), "bank_name": party_a.get("开户行", "")},
        "period": {"start": start.isoformat(), "end": end.isoformat(),
                   "period_cn": f"{start.year} 年 {start.month} 月", "asof": asof_day.isoformat(),
                   "asof_cn": f"{asof_day.year} 年 {asof_day.month:02d} 月 {asof_day.day:02d} 日",
                   "pay_cn": (f"{parse_day(pay).year} 年 {parse_day(pay).month:02d} 月 {parse_day(pay).day:02d} 日"
                              if pay else "主档未登记，以合同原文为准"),
                   "next_month": asof_day.month, "out_name": out_name},
        "sources": sources,
        "original_sales": _load_rows(args.original_sales, "历史销售"),
        "original_receipts": _load_rows(args.original_receipts, "历史签收"),
    }
    if mapping is not None:
        cfg["mapping"] = mapping
    if args.out:
        cfg["out_file"] = os.path.abspath(args.out)
    if args.overwrite:
        cfg["overwrite"] = True
    if business_type == "consignment":
        koujing_path = args.koujing or os.path.join(root, koujing.FILENAME)
        if args.allow_legacy:
            record = koujing.blank(args.scene, args.period, asof_day.isoformat())
            record["confirmed_by"] = "legacy"
            record["config_sha256"] = koujing.digest(record)
            print("[警告] --allow-legacy 仅供历史铺货只读复核，不作为正式交付件。", file=sys.stderr)
        else:
            record = koujing.load(koujing_path, args.scene)
        if args.mode == "incremental" and not args.in_workbook:
            _fail("[参数错误] incremental 需要 --in-workbook")
        cfg["koujing"] = record
        cfg["v17"] = {"enabled": True, "mode": args.mode,
                      "input_workbook": os.path.abspath(args.in_workbook) if args.in_workbook else "",
                      "pdf": bool(args.pdf), "pdf_sheets": [x.strip() for x in args.pdf_sheets.split(",") if x.strip()],
                      "koujing_path": os.path.abspath(koujing_path)}
    else:
        cfg["v17"] = {"enabled": False}
    payload = json.dumps(cfg, ensure_ascii=False)
    if args.cfg_out:
        cfg_out = os.path.abspath(args.cfg_out)
        try:
            os.makedirs(os.path.dirname(cfg_out), exist_ok=True)
            temp = cfg_out + ".tmp"
            with open(temp, "w", encoding="utf-8") as handle:
                handle.write(payload + "\n")
            os.replace(temp, cfg_out)
        except OSError as exc:
            _fail(f"[配置写入失败] {cfg_out}：{exc}")
    print(payload)


if __name__ == "__main__":
    main()
