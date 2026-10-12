"""End-to-end smoke test for Resume-Matcher with DeepSeek (corrected API shapes).

Response envelope is {request_id, data}. Resume detail is a QUERY param:
  GET /resumes?resume_id=X  -> data.raw_resume.processing_status / data.processed_resume
"""

import json
import sys
import time

import requests

BASE = "http://127.0.0.1:8000/api/v1"
JD_PATH = r"D:\dpsk\samples\jd-backend-ai.txt"
OUT = r"D:\dpsk\samples\improve-preview-result.json"

RESUME_ID = "c49b57a9-5296-4fbd-98ff-5b542c39d4a9"  # already uploaded + parsed


def log(step, msg):
    print(f"\n[{step}] {msg}", flush=True)


def get_resume(rid):
    r = requests.get(f"{BASE}/resumes", params={"resume_id": rid}, timeout=60)
    r.raise_for_status()
    return r.json()["data"]


def main():
    # ---- 0. confirm the parsed resume ----
    log("0", f"读取已解析简历 {RESUME_ID}")
    data = get_resume(RESUME_ID)
    raw = data["raw_resume"]
    parsed = data["processed_resume"]
    print(f"    processing_status = {raw['processing_status']}")
    pi = parsed["personalInfo"]
    print(f"    姓名: {pi['name']} | 职位: {pi['title']} | 城市: {pi['location']}")
    print(f"    工作经历: {len(parsed['workExperience'])} 段")
    print(f"    项目经历: {len(parsed['personalProjects'])} 个")
    print(f"    教育: {len(parsed['education'])} 条")
    print(f"    技能字段长度: {len(parsed['additional']['technicalSkills'])} 字符")

    # ---- 1. upload JD ----
    log("1", "上传职位描述")
    with open(JD_PATH, encoding="utf-8") as fh:
        jd_text = fh.read()
    j = requests.post(
        f"{BASE}/jobs/upload",
        json={"job_descriptions": [jd_text], "resume_id": RESUME_ID},
        timeout=120,
    )
    print(f"    HTTP {j.status_code}")
    if j.status_code >= 400:
        print(j.text[:2000])
        sys.exit(1)
    jr = j.json()
    print(f"    keys = {list(jr.keys())}")
    jd_data = jr.get("data", jr)
    print(f"    data keys = {list(jd_data.keys()) if isinstance(jd_data, dict) else type(jd_data)}")
    job_ids = jd_data.get("job_id") or jd_data.get("job_ids") or []
    if isinstance(job_ids, str):
        job_ids = [job_ids]
    job_id = job_ids[0] if job_ids else None
    print(f"    job_id = {job_id}")
    if not job_id:
        print(json.dumps(jr, ensure_ascii=False)[:2000])
        sys.exit(1)

    # ---- 2. core agent pipeline ----
    log("2", "核心流水线 /resumes/improve/preview")
    print("    链路: 提取JD关键词 -> 技能目标规划 -> 生成diff -> 应用diff -> 校验diff -> 润色")
    t0 = time.time()
    p = requests.post(
        f"{BASE}/resumes/improve/preview",
        json={"resume_id": RESUME_ID, "job_id": job_id},
        timeout=900,
    )
    dt = time.time() - t0
    print(f"    HTTP {p.status_code}   耗时 {dt:.1f}s")
    if p.status_code >= 400:
        print(p.text[:3000])
        sys.exit(1)

    pr = p.json()
    body = pr.get("data", pr)
    print(f"    顶层 keys = {list(pr.keys())}")
    print(f"    data keys = {list(body.keys()) if isinstance(body, dict) else type(body)}")

    with open(OUT, "w", encoding="utf-8") as fh:
        json.dump(pr, fh, ensure_ascii=False, indent=2)

    imps = body.get("improvements") or []
    print(f"\n    === 改进项 {len(imps)} 条 ===")
    for it in imps[:30]:
        if isinstance(it, dict):
            path = it.get("path") or it.get("field") or ""
            reason = (it.get("reason") or it.get("message") or it.get("rationale") or "")
            print(f"      - {path}")
            if reason:
                print(f"        {reason[:150]}")

    for key in ("keywords", "missing_keywords", "keyword_matches", "skills", "warnings"):
        val = body.get(key)
        if val:
            print(f"\n    === {key} ===")
            if isinstance(val, list):
                for v in val[:40]:
                    print(f"      - {v}")
            else:
                print(f"      {val}")

    print(f"\n    完整结果: {OUT}")
    print("\n[OK] 端到端跑通")


if __name__ == "__main__":
    main()
