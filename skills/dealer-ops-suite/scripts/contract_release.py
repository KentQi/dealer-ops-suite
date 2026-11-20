#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""合同包写前审阅、最终放行凭证与三文件哈希校验；不替代实质条款判断。"""
import hashlib
import json
import os
from datetime import datetime

SCHEMA = "dealer-ops/contract-release@1"
REVIEW_SCHEMA = "dealer-ops/contract-review@1"
ALLOWED_RESOLUTIONS = {"已修正", "接受风险", "不适用"}


def sha256_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def findings_digest(findings):
    payload = json.dumps(findings, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def load_resolutions(path):
    if not path:
        return {}
    with open(path, encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError("处理记录须为 {finding_id: {status,note,owner}} 对象")
    return data


def final_files(main_file, fc_file, policy_file):
    files = {}
    for role, label, path in (("main", "主合同", main_file), ("anti_diversion", "防窜货附件", fc_file),
                              ("price_policy", "最终价格政策", policy_file)):
        if not path or not os.path.isfile(path):
            raise ValueError(f"放行必须提供可读取的{label}最终文件")
        files[role] = {"path": os.path.abspath(path), "sha256": sha256_file(path)}
    if len({item["path"] for item in files.values()}) != 3:
        raise ValueError("主合同、防窜货附件、价格政策必须是三份不同文件")
    return files


def _review_problems(findings, llm_review):
    problems = []
    if not isinstance(findings, list) or not all(isinstance(item, dict) for item in findings):
        return ["审阅 findings 必须是对象数组"]
    if not isinstance(llm_review, dict) or llm_review.get("status") != "merged":
        problems.append("LLM 语义审阅尚未合并；不得把确定性检查当成完整合同审阅")
    for item in findings:
        if item.get("severity") == "blocker":
            problems.append(f"仍有阻断项 {item.get('id', '?')}：{item.get('message', '')}")
    return problems


def _clean_resolutions(findings, resolutions):
    problems = []
    known = {item.get("id") for item in findings if item.get("severity") == "high"}
    all_ids = {item.get("id") for item in findings}
    cleaned = {}
    for item_id in sorted(known):
        resolution = resolutions.get(item_id)
        if not isinstance(resolution, dict):
            problems.append(f"高风险项 {item_id} 尚无处理记录")
            continue
        if resolution.get("status") not in ALLOWED_RESOLUTIONS:
            problems.append(f"{item_id} 处理状态须为已修正／接受风险／不适用")
        if not str(resolution.get("note") or "").strip():
            problems.append(f"{item_id} 缺少处理说明")
        if not str(resolution.get("owner") or "").strip():
            problems.append(f"{item_id} 缺少责任人")
        cleaned[item_id] = {"status": resolution.get("status"),
                            "note": str(resolution.get("note") or "").strip(),
                            "owner": str(resolution.get("owner") or "").strip()}
    unknown = sorted(set(resolutions) - all_ids)
    if unknown:
        problems.append("处理记录含不存在的 finding id：" + "、".join(str(item) for item in unknown))
    return cleaned, problems


def _stored_resolution_problems(findings, stored):
    if not isinstance(stored, dict):
        return ["放行凭证缺 resolved_high"]
    expected = {item.get("id") for item in findings if item.get("severity") == "high"}
    problems = []
    if set(stored) != expected:
        missing = sorted(expected - set(stored))
        extra = sorted(set(stored) - expected)
        if missing:
            problems.append("高风险处理记录缺失：" + "、".join(str(item) for item in missing))
        if extra:
            problems.append("高风险处理记录含无关 id：" + "、".join(str(item) for item in extra))
    for item_id in sorted(expected & set(stored)):
        resolution = stored[item_id]
        if not isinstance(resolution, dict):
            problems.append(f"{item_id} 处理记录不是对象")
            continue
        if resolution.get("status") not in ALLOWED_RESOLUTIONS:
            problems.append(f"{item_id} 处理状态无效")
        if not str(resolution.get("note") or "").strip() or not str(resolution.get("owner") or "").strip():
            problems.append(f"{item_id} 处理说明或责任人为空")
    return problems


INPUT_ROLES = ("spec", "values", "fc_spec", "fc_values", "policy")


def input_files(paths):
    """把审阅输入绑定到规范化路径与内容哈希，防止换值或跨合同复用凭证。"""
    if not isinstance(paths, dict) or any(not paths.get(role) for role in INPUT_ROLES):
        raise ValueError("写前审阅必须绑定主合同／附件 spec、values 与价格政策文件")
    files = {}
    for role in INPUT_ROLES:
        path = os.path.abspath(paths[role])
        if not os.path.isfile(path):
            raise ValueError(f"审阅输入文件不存在（{role}）：{path}")
        files[role] = {"path": path, "sha256": sha256_file(path)}
    return files


def _input_problems(document, release, expected_inputs=None, required=True):
    files = release.get("input_files")
    if not isinstance(files, dict):
        return ["审阅凭证缺少输入文件哈希绑定"] if required else []
    problems = []
    declared = document.get("inputs") if isinstance(document.get("inputs"), dict) else {}
    for role in INPUT_ROLES:
        entry = files.get(role)
        if not isinstance(entry, dict) or not entry.get("path") or not entry.get("sha256"):
            problems.append(f"审阅输入凭证缺字段：{role}")
            continue
        path = entry["path"]
        if declared.get(role) and os.path.realpath(str(declared[role])) != os.path.realpath(path):
            problems.append(f"审阅记录 inputs.{role} 与凭证绑定文件不一致")
        if expected_inputs and expected_inputs.get(role):
            if os.path.realpath(os.path.abspath(expected_inputs[role])) != os.path.realpath(path):
                problems.append(f"本次 --{role.replace('_', '-')} 不是写前审阅绑定的文件")
        if not os.path.isfile(path):
            problems.append(f"审阅输入文件不存在（{role}）：{path}")
            continue
        try:
            if sha256_file(path) != entry["sha256"]:
                problems.append(f"审阅输入文件在凭证生成后已变更（{role}）")
        except OSError as exc:
            problems.append(f"审阅输入文件不可读（{role}）：{exc}")
    return problems


def build_release(findings, llm_review, main_file, fc_file, policy_file, resolutions_path, inputs=None):
    problems = _review_problems(findings, llm_review)
    resolved, resolution_problems = _clean_resolutions(findings, load_resolutions(resolutions_path))
    problems += resolution_problems
    if problems:
        raise ValueError("；".join(problems))
    if inputs is None:
        # 放行凭证必须绑定写前审阅的 values/spec/policy 等输入：
        # 否则理论上可用未经写前审阅的 values 生成放行。
        raise ValueError("最终放行必须绑定写前审阅输入（--spec/--values/--fc-spec/--fc-values/--policy）")
    document = {"schema": SCHEMA, "status": "released", "released_at": datetime.now().isoformat(timespec="seconds"),
                "files": final_files(main_file, fc_file, policy_file), "resolved_high": resolved,
                "review_findings_digest": findings_digest(findings)}
    document["input_files"] = input_files(inputs)
    return document


def build_write_approval(findings, llm_review, resolutions_path, inputs=None):
    """只形成写前结构预检凭证，不要求尚未生成的成品文件或哈希。"""
    problems = _review_problems(findings, llm_review)
    resolved, resolution_problems = _clean_resolutions(findings, load_resolutions(resolutions_path))
    problems += resolution_problems
    if problems:
        raise ValueError("；".join(problems))
    return {"schema": SCHEMA, "status": "write_approved", "resolved_high": resolved,
            "input_files": input_files(inputs), "review_findings_digest": findings_digest(findings)}


def check_write_review(document, expected_inputs=None):
    """写前门禁：绑定 findings、高风险处理和本次实际审阅输入。"""
    if not isinstance(document, dict) or document.get("schema") != REVIEW_SCHEMA:
        return ["写前审阅记录 schema 无效"]
    findings = document.get("findings")
    problems = _review_problems(findings, document.get("llm_review"))
    if problems:
        return problems
    release = document.get("release")
    if not isinstance(release, dict) or release.get("schema") != SCHEMA or release.get("status") not in (
            "write_approved", "released"):
        return problems + ["没有有效的合同包写前审阅凭证"]
    if release.get("review_findings_digest") != findings_digest(findings):
        problems.append("审阅 findings 在写前凭证生成后已变更")
    problems += _stored_resolution_problems(findings, release.get("resolved_high"))
    problems += _input_problems(document, release, expected_inputs)
    return problems


def verify_release(document):
    """只认最终 released 对象，并重新读取三份最终文件计算哈希。"""
    if not isinstance(document, dict) or document.get("schema") != REVIEW_SCHEMA:
        return False, ["审阅记录 schema 无效"]
    release = document.get("release")
    if not isinstance(release, dict) or release.get("schema") != SCHEMA or release.get("status") != "released":
        return False, ["没有有效合同包最终放行凭证"]
    findings = document.get("findings")
    problems = _review_problems(findings, document.get("llm_review"))
    if problems:
        return False, problems
    if release.get("review_findings_digest") != findings_digest(findings):
        problems.append("审阅 findings 在放行后已变更")
    problems += _stored_resolution_problems(findings, release.get("resolved_high"))
    problems += _input_problems(document, release, required=True)
    files = release.get("files")
    if not isinstance(files, dict):
        return False, problems + ["放行凭证缺最终文件清单"]
    for role in ("main", "anti_diversion", "price_policy"):
        entry = files.get(role) or {}
        path = entry.get("path")
        if not path or not os.path.isfile(path):
            problems.append(f"最终文件不存在（{role}）：{path}")
            continue
        try:
            if sha256_file(path) != entry.get("sha256"):
                problems.append(f"最终文件在放行后已变更（{role}）：{path}")
        except OSError as exc:
            problems.append(f"最终文件不可读（{role}）：{exc}")
    return not problems, problems


def verify_final_release(document, result_path):
    """最终验收除三文件哈希外，还要求当前主合同就是放行时那份文件。"""
    ok, problems = verify_release(document)
    if isinstance(document, dict):
        main = ((document.get("release") or {}).get("files") or {}).get("main") or {}
        main_path = main.get("path")
        if not main_path or not result_path or os.path.realpath(main_path) != os.path.realpath(result_path):
            problems.append("当前 --result 不是放行凭证绑定的最终主合同")
            ok = False
    return ok and not problems, problems
