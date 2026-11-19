#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
plan_trust.py —— 计划可信性公共模块（只读）

为什么需要它（第三轮对抗性审计的结论）：

  plan.json 是普通 JSON，可以被编辑。只比对 op["value"] 与 op["expect"] 没有意义——
  两者都在同一份文件里，一起改就能绕过；values_digest 也只是 plan 里记录的一个
  字符串，改 ops 时原样保留即可。实测：把 JX-18 的 value 与 expect 同时改成
  「季度返利：0%」，并把成品的对应文字同步改掉，旧版验收仍判「✓ 通过」。

  唯一可靠的做法是：拿「用户确认值 values.json + 规格 spec + 原始模板」用
  build_plan.py **重新编译**一份参考计划，逐条核对 plan.ops。
  这样验收证明的是「成品 == 计划 == 用户确认值」，而不只是「成品 == 计划」。

  同时校验 spec 文件内容指纹，防止 --spec 指向一份被放宽（必填改 false、
  readonly_check 清空、render 被改）的伪造规格。

信任根：values.json 与 scripts/spec/*.json。它们必须由用户确认 / 由 skill 维护，
本模块只负责证明「成品与计划确实由这两者派生」，不负责证明这两者自身没被篡改。
"""
import hashlib
import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
BUILD_PLAN = os.path.join(HERE, "build_plan.py")


def file_sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def load_values(path):
    raw = json.load(open(path, encoding="utf-8"))
    return {k: v for k, v in raw.items() if not k.startswith("_")}


def values_digest(values):
    return hashlib.sha256(
        json.dumps(values, ensure_ascii=False, sort_keys=True).encode("utf-8")
    ).hexdigest()


def canon(x):
    """归一成可精确比较的串（忽略 dict 键顺序，不忽略任何字段）。"""
    return json.dumps(x, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def recompile(spec_path, base_doc, values_path, workdir):
    """调用 build_plan.py 重编译，返回 (returncode, ref_plan, log)。"""
    out = os.path.join(workdir, "ref-plan.json")
    proc = subprocess.run(
        [sys.executable, BUILD_PLAN,
         "--spec", os.path.abspath(spec_path),
         "--doc", os.path.abspath(base_doc),
         "--values", os.path.abspath(values_path),
         "--out", out],
        capture_output=True, text=True, timeout=180,
    )
    path = out if os.path.exists(out) else out + ".rejected"
    ref = json.load(open(path, encoding="utf-8")) if os.path.exists(path) else None
    return proc.returncode, ref, (proc.stdout + proc.stderr).strip()


def check_plan(plan, spec_path, values_path, base_doc):
    """核对一份计划是否真的由「这份 spec + 这份 values + 这个基准文档」编译而来。

    返回 problems 列表；空列表 = 可信。
    """
    problems = []
    if not spec_path or not values_path:
        return ["缺少 --spec / --values：无法核实计划是否由用户确认值编译而来，已拒绝"]
    if not base_doc or not os.path.exists(base_doc):
        return [f"重编译基准文档不存在：{base_doc}"]

    # 1) 值指纹 + 计划内记录的原始确认值，必须与当前 values.json 逐字一致
    try:
        values = load_values(values_path)
    except (OSError, ValueError) as e:
        return [f"无法读取确认值文件：{e}"]
    if not plan.get("values_digest"):
        problems.append("计划缺少值指纹（旧版或手工计划），请重跑 build_plan")
    elif values_digest(values) != plan["values_digest"]:
        problems.append("值指纹不符：改了确认值却没重编译计划，请重跑 build_plan")
    if "source_values" not in plan:
        problems.append("计划缺少 source_values，无法核对确认值来源，请重跑 build_plan")
    elif canon(plan["source_values"]) != canon(values):
        problems.append("计划内记录的确认值与当前 values.json 不一致，请重跑 build_plan")

    # 2) 规格必须是编译时的那一份（内容指纹，不是路径）
    if "spec_sha256" not in plan:
        problems.append("计划缺少规格指纹（旧版计划），请用当前版 build_plan 重新编译")
    else:
        try:
            if file_sha256(spec_path) != plan["spec_sha256"]:
                problems.append("规格文件与编译时不一致：--spec 指向了另一份 spec，已拒绝")
        except OSError as e:
            problems.append(f"无法读取规格文件：{e}")

    # 3) 重新编译，逐条核对 ops —— 这一步堵死「value 与 expect 一起改」
    try:
        with tempfile.TemporaryDirectory(prefix="cfj-trust-") as wd:
            rc, ref, log = recompile(spec_path, base_doc, values_path, wd)
    except (OSError, subprocess.TimeoutExpired, ValueError) as e:
        return problems + [f"重编译失败：{e}"]
    if ref is None:
        return problems + [f"无法重编译参考计划：{log[-300:]}"]
    if rc != 0:
        problems.append("按当前 spec + values 重编译存在错误，计划不可执行：" + log[-300:])

    a, b = plan.get("ops", []), ref.get("ops", [])
    if len(a) != len(b):
        problems.append(
            f"计划操作数与重编译结果不符（计划 {len(a)} 条 / 应为 {len(b)} 条），"
            "计划可能被手工增删过"
        )
    else:
        shown = 0
        for i, (x, y) in enumerate(zip(a, b)):
            if canon(x) == canon(y):
                continue
            shown += 1
            if shown <= 8:
                problems.append(
                    f"第 {i + 1} 条操作与用户确认值的渲染结果不符："
                    f"计划「{x.get('field_id', '?')}={str(x.get('value', ''))[:26]}」，"
                    f"应为「{str(y.get('value', ''))[:26]}」"
                )
        if shown > 8:
            problems.append(f"…… 另有 {shown - 8} 条操作不符")
    return problems


def check_plan_selfconsistency(plan):
    """计划内部自洽：每条 op 必须带 expect 且与 value 精确相等。"""
    problems = []
    for op in plan.get("ops", []):
        exp = op.get("expect")
        if exp is None:
            problems.append(f"{op.get('field_id', '?')}：计划缺少 expect，无法核实落笔内容")
            continue
        val = op.get("value")
        # 精确比较：str() 归一会把 1 与 "1"、True 与 "True" 判成一致，
        # 让伪造计划借类型混淆蒙混过关。
        if exp != val or type(exp) is not type(val):
            problems.append(
                f"{op.get('field_id', '?')}：落笔文本与 expect 不一致"
                f"（应为「{str(exp)[:30]}」，实为「{str(val)[:30]}」）"
            )
    return problems
