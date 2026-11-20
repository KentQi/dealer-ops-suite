#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
apply_fill.py —— 按「已确认填充计划」批量写入合同（经 editor_sdk，可可视化）

设计约束（与 SKILL.md 的硬边界一一对应）：
  · 只写计划里列出的位置，不新增/删除段落、不改样式、不动条款文字；
  · 全部定位先解析、再按文档位置从后往前写，避免坐标漂移；
  · 任一操作定位失败即整单中止（已写入部分由调用方用备份还原）；
  · 写入内容逐条回显，便于事后核对。

用法：
  python3 apply_fill.py --plan plan.json --edsdk <editor_sdk>/edsdk.py [--dry-run]

plan.json 结构：
{
  "doc": "JX",
  "file_id": "<可选，缺省读 --file-id>",
  "ops": [
    {"field_id":"JX-01","kind":"insert","find":"主合同编号：","occurrence":1,"value":"JX-2026-001"},
    {"field_id":"JX-02","kind":"replace","find":"＿＿＿＿＿＿＿＿＿＿＿＿","occurrence":1,"value":"某某有限公司"},
    {"field_id":"JX-C1","kind":"cell","table":2,"row":0,"col":0,"value":"第一季度：建立并有效运营5"}
  ]
}
  kind=insert  → 在 find 匹配结束处插入 value
  kind=replace → 用 value 替换 find 匹配文本
  kind=cell    → 整体改写单元格文本（table/row/col 为 0 基，脚本内部转 1 基）
"""
import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile

PY = sys.executable

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from plan_trust import check_plan, check_plan_selfconsistency, load_values, canon  # noqa: E402


class EDSDKError(RuntimeError):
    pass


def check_incremental_plan(plan, spec_path, values_path):
    """增量修复计划（无值指纹）写入前的信任根绑定核对。

    增量 plan 是手工从全量计划摘出的 op 子集，没有 values_digest 可供重编译复核，
    因此必须显式绑定「用户确认值 + 规格」：每个 op 的字段必须是规格内已确认字段，
    落笔文本必须与确认值（或其渲染结果）逐字一致；否则不得写入。
    """
    problems = check_plan_selfconsistency(plan)
    try:
        with open(spec_path, encoding="utf-8") as fh:
            spec = json.load(fh)
    except (OSError, ValueError) as exc:
        return problems + [f"无法读取规格文件：{exc}"]
    try:
        values = load_values(values_path)
    except (OSError, ValueError) as exc:
        return problems + [f"无法读取确认值文件：{exc}"]
    known_ids = {f["id"] for f in spec.get("fields", [])}
    known_ids |= {c["id"] for c in spec.get("cells", [])}
    known_ids |= {g["id"] for g in spec.get("grids", [])}
    spec_fields = {f["id"]: f for f in spec.get("fields", [])}
    for op in plan.get("ops", []):
        fid = op.get("field_id", "?")
        if known_ids and fid not in known_ids:
            problems.append(f"{fid}：不在规格字段清单中，增量计划不得引入规格外字段")
            continue
        if fid not in values:
            problems.append(f"{fid}：values.json 无该字段的用户确认值；"
                            f"先把确认值补进 values.json 再执行增量修复")
            continue
        if op.get("kind") == "cell":
            continue  # 表格格值由全量验收兜底，这里只要求字段已确认
        field = spec_fields.get(fid) or {}
        raw = values[fid]
        rendered = raw
        if isinstance(raw, str) and field.get("render"):
            from build_plan import render
            rendered = render(field["render"], raw.strip(), field.get("vtype"))
        if canon(op.get("value")) not in (canon(raw), canon(rendered)):
            problems.append(f"{fid}：落笔文本与用户确认值不符（应为「{str(rendered)[:30]}」，"
                            f"实为「{str(op.get('value'))[:30]}」），已拒绝")
    return problems


def same_file(a, b):
    """路径等价：realpath 相同，或两者是同一 inode（硬链接 / 绑定挂载）。

    只比 realpath 会被硬链接绕过：模板与输出可以是两个不同路径、却指向同一 inode，
    写入会直接毁掉模板。第四轮集成审计已实证。
    """
    if not a or not b:
        return False
    if os.path.realpath(a) == os.path.realpath(b):
        return True
    try:
        sa, sb = os.stat(a), os.stat(b)
        return (sa.st_dev, sa.st_ino) == (sb.st_dev, sb.st_ino)
    except OSError:
        return False


def content_hash(path):
    try:
        h = hashlib.sha256()
        with open(path, "rb") as fh:
            for chunk in iter(lambda: fh.read(1 << 20), b""):
                h.update(chunk)
        return h.hexdigest()
    except OSError:
        return None


def sample_needles(ops, limit=5):
    """从计划里取几条写入值做「保存后读回」抽样。"""
    out = []
    for op in ops:
        clean = re.sub(r"\s+", "", str(op.get("value", "")))[:12]
        if len(clean) >= 4:
            out.append(clean)
        if len(out) >= limit:
            break
    return out


def readback_check(path, needles):
    """保存后读回：必须是有效 DOCX，且抽样写入值确实出现在 document.xml 里。

    只比 mtime/size 会被「touch」「写等长垃圾」「保存到别的路径」骗过。
    """
    import html
    if not zipfile.is_zipfile(path):
        return False, "保存后的文件不是有效的 DOCX（ZIP）包"
    try:
        with zipfile.ZipFile(path) as z:
            s = z.read("word/document.xml").decode("utf-8", "ignore")
    except (KeyError, OSError, zipfile.BadZipFile) as e:
        return False, f"无法读取 word/document.xml：{e}"
    text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", s))
    compact = re.sub(r"\s+", "", html.unescape(text))
    miss = [n for n in needles if n and n not in compact]
    if miss:
        return False, (f"写入内容未能在保存后的文件中读回（{len(miss)}/{len(needles)} 项抽样缺失，"
                       f"例：「{miss[0][:20]}」）")
    return True, ""


def call(edsdk, tool, file_id, args, json_required=True):
    payload = dict(args)
    if file_id:
        payload["file_id"] = file_id
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False,
                                     encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False)
        path = fh.name
    try:
        proc = subprocess.run(
            [PY, edsdk, "call", tool, "--json-file", path],
            capture_output=True, text=True, timeout=180,
        )
    finally:
        os.unlink(path)
    if proc.returncode != 0:
        raise EDSDKError(f"{tool} 调用失败：{proc.stderr.strip()[:400]}")
    out = proc.stdout.strip()
    if not out:
        if json_required:
            raise EDSDKError(f"{tool} 无返回")
        return {}
    try:
        data = json.loads(out)
    except json.JSONDecodeError:
        if json_required:
            raise EDSDKError(f"{tool} 返回非 JSON：{out[:300]}")
        return {"raw": out}
    if isinstance(data, dict):
        if data.get("error") or data.get("isError"):
            raise EDSDKError(
                f"{tool} 报错：{data.get('error') or data.get('message') or str(data)[:300]}")
        if data.get("success") is False:
            raise EDSDKError(f"{tool} 返回失败：{str(data)[:300]}")
    return data


def resolve(edsdk, file_id, op, table_ids):
    """把一条 op 解析成 {pos, action}；定位不唯一或找不到直接抛错。"""
    kind = op["kind"]
    if kind == "cell":
        info = table_ids.get(op["table"])
        if not info:
            raise EDSDKError(f"{op['field_id']}：文档里没有第 {op['table'] + 1} 张表")
        return {"op": op, "pos": info.get("pos", 0), "kind": "cell",
                "table_id": info.get("id"), "table_pos": info.get("pos", 0)}

    find = op["find"]
    occ = int(op.get("occurrence", 1))
    data = call(edsdk, "doc_find", file_id, {"text": find})
    locs = sorted(data.get("locations", []), key=lambda x: x["begin"])
    if not locs:
        raise EDSDKError(f"{op['field_id']}：文档中找不到「{find}」")
    if len(locs) < occ:
        raise EDSDKError(
            f"{op['field_id']}：需要第 {occ} 处「{find}」，全文只有 {len(locs)} 处"
        )
    hit = locs[occ - 1]

    # 上下文校验：计划里带了 ctx 就必须命中，防止模板改版后写错位置
    ctx = op.get("ctx")
    if ctx:
        flat = hit.get("related_text", "").replace("[", "").replace("]", "")
        probe = ctx.replace("[", "").replace("]", "")
        if probe not in flat:
            raise EDSDKError(
                f"{op['field_id']}：上下文校验不通过\n"
                f"    期望包含：{probe[:60]}\n    实际：{flat[:60]}"
            )

    # 幂等保护：追加类操作若发现「锚点后面已经就是这个值」，说明计划执行过，
    # 继续写会变成「联系人：张三张三」。
    # 旧版只看 related_text 里 "]"+值前 20 字，实测可被「去掉方括号」「中间插全角空格」
    # 「related_text 窗口截断」「锚点第 2 处」绕过，故改为去空白的整串比对 + 全文查找兜底。
    if kind == "insert":
        val = str(op["value"])
        needle = str(op["find"]) + val
        flat = hit.get("related_text", "").replace("[", "").replace("]", "")
        if re.sub(r"\s+", "", needle) in re.sub(r"\s+", "", flat):
            raise EDSDKError(
                f"{op['field_id']}：锚点后已存在「{val}」，该计划看起来已经执行过，"
                f"重复写入会叠字，已中止"
            )
        probe = call(edsdk, "doc_find", file_id, {"text": needle})
        if not isinstance(probe, dict):
            raise EDSDKError(f"{op['field_id']}：幂等复核返回结构异常，已中止")
        if probe.get("locations"):
            raise EDSDKError(
                f"{op['field_id']}：全文已存在「{str(op['find'])}{val}」，"
                f"该计划看起来已经执行过，重复写入会叠字，已中止"
            )

    return {"op": op, "pos": hit["begin"], "kind": kind,
            "begin": hit["begin"], "end": hit["end"], "related": hit.get("related_text", "")}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--plan", required=True)
    ap.add_argument("--edsdk", required=True)
    ap.add_argument("--file-id")
    ap.add_argument("--protect", required=True,
                    help="受保护路径（原始模板）：必需；目标正是它则拒绝写入")
    ap.add_argument("--spec", required=True,
                    help="字段规格 JSON（写入前信任根之一；会与写前审阅凭证绑定核对）")
    ap.add_argument("--values", required=True,
                    help="用户确认值 JSON（写入前信任根之一；增量计划逐条核对落笔值）")
    ap.add_argument("--review", required=True,
                    help="review_contract_package.py --approve-write 生成的写前审阅 JSON")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--no-save", action="store_true")
    args = ap.parse_args()

    from contract_release import check_write_review
    try:
        with open(args.review, encoding="utf-8") as fh:
            review = json.load(fh)
    except (OSError, ValueError) as exc:
        sys.exit(f"无法读取写前审阅记录：{exc}")
    review_problems = check_write_review(review, {"spec": args.spec, "values": args.values})
    if review_problems:
        sys.exit("写前审阅门禁未通过，未写入任何内容：" + "；".join(review_problems))
    print("写前审阅门禁通过：LLM 已合并、无阻断项，高风险处理记录完整。")

    if args.plan.endswith(".rejected"):
        sys.exit("这是 build_plan 校验失败时留存的排查文件（.rejected），不是可执行计划，已拒绝")
    with open(args.plan, encoding="utf-8") as fh:
        plan = json.load(fh)
    file_id = args.file_id or plan.get("file_id")
    if not file_id:
        sys.exit("缺少 file_id：请传 --file-id 或写进 plan.json")

    # 写入前可信复核：全量计划必须能由「spec + values + 模板」重编译复现。
    # 第三轮审计 PoC：只改 plan 的 value+expect 并同步改成品，旧版验收照样放行；
    # 这里把同一道闸门提前到落笔之前，避免污染成品。
    if "values_digest" in plan:
        if not args.spec or not args.values:
            sys.exit("这份计划由 build_plan 生成，写入前必须同时给 --spec 与 --values "
                     "做可信复核（不得手工编辑 plan.json）")
        base = plan.get("template")
        if not base or not os.path.exists(base):
            sys.exit(f"计划记录的原始模板不存在：{base}，无法重编译复核")
        problems = check_plan(plan, args.spec, args.values, base)
        if problems:
            print("✗ 计划不可信，未写入任何内容：")
            for p in problems[:10]:
                print("   ·", p)
            sys.exit(7)
        print(f"计划可信性复核通过：{len(plan.get('ops', []))} 条操作与确认值逐条一致")
    else:
        # 增量修复计划（无值指纹）：不得静默写入——逐条绑定用户确认值与规格。
        inc_problems = check_incremental_plan(plan, args.spec, args.values)
        if inc_problems:
            print("✗ 增量计划未通过信任根核对，未写入任何内容：")
            for p in inc_problems[:10]:
                print("   ·", p)
            sys.exit(7)
        print("增量修复计划（无值指纹）：已逐条核对落笔值与用户确认值一致；"
              "收尾仍须用全量计划重跑 verify_fill.py，否则不得交付。")

    # file_id 只是内存实例的键，与路径没有强制绑定；写错实例会把内容灌进另一份文档。
    # 因此先查出这个 file_id 到底落在哪个路径，再和受保护路径比对。
    target_path = None
    try:
        pool = call(args.edsdk, "get_pool_status", None, {})
        if not isinstance(pool, dict):
            raise EDSDKError("get_pool_status 返回的不是对象")
        editors = pool.get("open_editors") or pool.get("editors") or []
        if isinstance(editors, dict):
            editors = list(editors.values())
        if not isinstance(editors, list):
            raise EDSDKError("get_pool_status 的 editors 结构异常")
        for ed in editors:
            if isinstance(ed, dict) and ed.get("file_id") == file_id:
                target_path = ed.get("file_path")
                break
    except (EDSDKError, AttributeError, TypeError) as e:
        # 身份无法确认时一律中止，不再依赖「有没有传 --protect」。
        sys.exit(f"无法确认目标文件身份，已中止：{e}")

    # 模板写保护不再依赖可选参数：--protect 已是必需参数，且 plan.template
    # 也一并纳入保护集合。用 same_file（realpath + inode）比对，堵住硬链接绕过。
    protect_paths = [p for p in (args.protect, plan.get("template")) if p]
    if not target_path:
        sys.exit(f"file_id「{file_id}」不在编辑器池中，无法确认它对应哪个文件，已中止")
    for pp in protect_paths:
        if same_file(target_path, pp):
            sys.exit(f"拒绝写入：目标文件就是模板本身（{target_path}，与 {pp} 是同一文件）。"
                     f"请先另存为新文件。")
    print(f"目标确认：{target_path}")

    stat_before = None
    hash_before = None
    if target_path:
        try:
            st = os.stat(target_path)
            stat_before = (st.st_mtime_ns, st.st_size)
        except OSError:
            pass
        hash_before = content_hash(target_path)

    ops = plan.get("ops", [])
    if not ops:
        sys.exit("计划为空，未执行任何写操作")

    # 表格 id 与位置，供 cell 类操作定位
    table_ids = {}
    tdata = call(args.edsdk, "doc_list_tables", file_id, {})
    for ti, t in enumerate(tdata.get("tables", [])):
        tid = t.get("table_id")
        if not tid or str(tid).startswith("tbl:"):
            tid = None
        table_ids[ti] = {"id": tid, "pos": t.get("table_start", 0)}

    print(f"解析 {len(ops)} 条操作 …")
    resolved = []
    errors = []
    for op in ops:
        try:
            r = resolve(args.edsdk, file_id, op, table_ids)
            # 坐标合法性：伪造/错配的 doc_find 返回负数、非整数或 end<begin 时，
            # 旧版会直接把垃圾坐标传给写接口，且冲突检测也失效。
            if r["kind"] in ("insert", "replace"):
                b, e = r.get("begin"), r.get("end")
                if (not isinstance(b, int) or not isinstance(e, int)
                        or isinstance(b, bool) or isinstance(e, bool)
                        or b < 0 or e < b):
                    raise EDSDKError(
                        f"{op.get('field_id')}：坐标非法 begin={b} end={e}")
            resolved.append(r)
        except EDSDKError as e:
            errors.append(str(e))
        except KeyError as e:
            errors.append(f"{op.get('field_id')}：缺少字段 {e}")
    if errors:
        print("\n✗ 定位失败，未写入任何内容：")
        for e in errors:
            print("   ·", e)
        sys.exit(2)

    # 从文档后部往前写，保证尚未处理的位置不漂移
    resolved.sort(key=lambda r: r["pos"], reverse=True)

    # 重叠 / 同点冲突检查：两个替换区间相交，或插入点落在别人的替换区间内，一律中止
    spans = []
    for r in resolved:
        if r["kind"] == "replace":
            spans.append((r["begin"], r["end"], r["op"]["field_id"]))
    # 插入用的是 r["end"]（doc_insert_text 的 idx），不是 r["begin"]；
    # 旧版拿 begin 去做冲突检测，等于用错了坐标。
    inserts = [(r["end"], r["op"]["field_id"]) for r in resolved if r["kind"] == "insert"]
    conflicts = []
    spans.sort()
    for i in range(len(spans) - 1):
        if spans[i][1] > spans[i + 1][0]:
            conflicts.append(f"{spans[i][2]} 与 {spans[i + 1][2]} 的替换区间重叠")
    for pos, fid in inserts:
        for b, e, fid2 in spans:
            if b < pos < e:
                conflicts.append(f"{fid} 的插入点落在 {fid2} 的替换区间内")
    dup = {}
    for pos, _fid in inserts:
        dup[pos] = dup.get(pos, 0) + 1
    for pos, n in dup.items():
        if n > 1:
            conflicts.append(f"位置 {pos} 有 {n} 个插入操作，写入顺序不确定，请合并为一条")
    if conflicts:
        print("\n✗ 操作冲突，未写入任何内容：")
        for c in sorted(set(conflicts)):
            print("   ·", c)
        print("  请改用 kind=cell（按表格行列定位）或加长 find 文本消歧。")
        sys.exit(4)

    print("-" * 78)
    for r in resolved:
        op = r["op"]
        if r["kind"] == "cell":
            print(f"  [单元格] {op['field_id']:<7} 表#{op['table']} "
                  f"R{op['row'] + 1}C{op['col'] + 1} ← {op['value']}")
        elif r["kind"] == "insert":
            print(f"  [续写]   {op['field_id']:<7} 「{op['find'][:18]}」后 ← {op['value']}")
        else:
            print(f"  [替换]   {op['field_id']:<7} 「{op['find'][:18]}」→ {op['value']}")
    print("-" * 78)

    if args.dry_run:
        print("dry-run：仅解析，未写入。")
        return

    done = 0
    for r in resolved:
        op = r["op"]
        try:
            if r["kind"] == "cell":
                payload = {"cells": [{"row": op["row"] + 1, "col": op["col"] + 1,
                                      "text": op["value"]}]}
                if r["table_id"]:
                    payload["table_id"] = r["table_id"]
                else:
                    payload["idx"] = r["table_pos"]
                call(args.edsdk, "doc_set_table_cells", file_id, payload)
            elif r["kind"] == "insert":
                call(args.edsdk, "doc_insert_text", file_id,
                     {"idx": r["end"], "text": op["value"]})
            else:
                call(args.edsdk, "doc_replace_text", file_id,
                     {"ranges": [{"begin": r["begin"], "end": r["end"]}],
                      "text": op["value"]})
            done += 1
            print(f"  ✓ {op['field_id']}")
        except EDSDKError as e:
            print(f"\n✗ {op['field_id']} 写入失败：{e}")
            print(f"  已写入 {done}/{len(resolved)} 条；内存实例现在是脏的，禁止交付、禁止复用。")
            print("  必须按「回滚三步」处理：")
            print('    ① cp "<输出>.bak" "<输出>"                还原磁盘')
            print(f"    ② close_file file_id={file_id}            关掉脏内存实例")
            print("    ③ 重新 present_files 打开，再取新的 file_id（切勿沿用本轮 ID）")
            print("  回滚后重跑 build_plan.py 得到全新计划；收尾必须由 verify_fill.py 判定。")
            sys.exit(3)

    if not args.no_save:
        out = call(args.edsdk, "save_file", file_id, {}, json_required=False)
        raw = str(out.get("raw", "")) if isinstance(out, dict) else str(out)
        low = raw.lower()
        problems = []
        # 工具级失败可能藏在非 JSON 文本里；call() 的 isError 检查只覆盖 JSON 返回。
        if any(k in low for k in ("error", "fail", "失败", "错误", "denied", "exception")):
            problems.append(f"save_file 返回失败文本：{raw[:160]}")
        if isinstance(out, dict) and not out:
            problems.append("save_file 无回执（空返回）")
        m = re.search(r"saved\s+to\s*:\s*(.+)", raw, re.I)
        if m and target_path:
            saved_path = m.group(1).strip().strip('"').strip("'")
            if os.path.basename(saved_path) != os.path.basename(target_path):
                problems.append(f"保存回执指向另一个文件（{saved_path[:120]}）")
        if target_path:
            try:
                st_after = os.stat(target_path)
            except OSError:
                st_after = None
                problems.append("保存后目标文件不可访问")
            if stat_before is not None and st_after is not None:
                if (st_after.st_mtime_ns, st_after.st_size) == stat_before:
                    problems.append("目标文件 mtime / 大小没有变化")
            # 只比 mtime/size 会被 touch 或「写等长垃圾」骗过，必须比内容哈希。
            if hash_before and content_hash(target_path) == hash_before:
                problems.append("目标文件内容哈希没有变化（touch 不算保存成功）")
            ok_read, why = readback_check(target_path, sample_needles(ops))
            if not ok_read:
                problems.append(why)
        if problems:
            print("\n✗ 保存结果不可信，禁止交付：")
            for p in problems:
                print("   ·", p)
            print(f"  save_file 返回：{str(out)[:200]}")
            print("  按「回滚三步」还原后重来（还原磁盘 → close_file → 重新打开取新 file_id）。")
            sys.exit(6)
        print(f"\n已保存并读回核验通过。共写入 {done} 条。")
        print("提醒：apply_fill.py 只负责落笔，仍须用 verify_fill.py 收尾验收，退出码 0 才可交付。")
    else:
        print(f"\n已写入 {done} 条（--no-save：未保存、未做保存核验）。此文件不得交付。")


if __name__ == "__main__":
    try:
        main()
    except EDSDKError as e:
        print(f"\n✗ {e}")
        print("已中止。若已有写入：先用 .bak 覆盖还原，再 close_file 关闭该实例并重新打开——"
              "内存里的脏内容不会因为磁盘还原而消失，复用同一实例会把旧内容再写回去。")
        sys.exit(5)
    except KeyboardInterrupt:
        print("\n已中断。请检查文档状态并按需还原。")
        sys.exit(130)
