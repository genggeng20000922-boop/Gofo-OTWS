#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
链路 A（构建时拉取）：读取飞书多维表 -> 生成 data.json

双模式，自动选择：
  · OpenAPI 模式（CI / 生产）：存在 LARK_APP_ID + LARK_APP_SECRET 时启用。
    自行换取 tenant_access_token 调 REST 接口，凭证只从环境变量读取。
  · CLI 模式（本地调试）：无凭证时回退到已登录的 lark-cli。

本脚本与产物中不含任何密钥。

用法:
    python build_data.py            # 生成 data.json
    python build_data.py --probe    # 只打印关键指标，不写文件
"""

import json
import os
import shutil
import subprocess
import sys
import time
import collections
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone, timedelta

BASE_TOKEN = "TorLbanFNaFDussPBckcwdvvnhc"
TABLE_PROGRESS = "tbl6Z0fPeMdHw61W"   # 2、区域推广进度
TABLE_PLAN = "tbl3mOJhdyedSGZH"       # 1、推广计划
OUT_PATH = "data.json"
CST = timezone(timedelta(hours=8))

API_HOST = "https://open.feishu.cn/open-apis"


# ---------------------------------------------------------------- 取数后端

class OpenApiBackend:
    """直连飞书 OpenAPI，凭证仅来自环境变量。"""

    def __init__(self, app_id, app_secret):
        self.app_id = app_id
        self.app_secret = app_secret
        self._token = None
        self._expire_at = 0

    def _request(self, method, path, payload=None, token=None):
        """统一的 HTTP 请求，支持 GET / POST；429 与 5xx 走指数退避重试。"""
        url = API_HOST + path
        headers = {"Content-Type": "application/json; charset=utf-8"}
        if token:
            headers["Authorization"] = "Bearer " + token
        body = json.dumps(payload).encode("utf-8") if payload is not None else None
        req = urllib.request.Request(url, data=body, headers=headers, method=method)
        for attempt in range(4):
            try:
                with urllib.request.urlopen(req, timeout=30) as resp:
                    return json.loads(resp.read().decode("utf-8"))
            except urllib.error.HTTPError as e:
                detail = e.read().decode("utf-8", "replace")
                # 429 限流 / 5xx 服务端错误 -> 指数退避重试
                if (e.code == 429 or e.code >= 500) and attempt < 3:
                    time.sleep(2 ** attempt)
                    continue
                raise RuntimeError("HTTP %s: %s" % (e.code, detail[:400]))
            except urllib.error.URLError as e:
                if attempt < 3:
                    time.sleep(2 ** attempt)
                    continue
                raise RuntimeError("网络错误: %s" % e)

    @staticmethod
    def _fingerprint(v):
        """返回值的可诊断指纹：长度 + 首尾字符，不泄露完整内容。"""
        if v is None:
            return "<None>"
        return "len=%d head=%r tail=%r has_space=%s" % (
            len(v), v[:8], v[-4:] if len(v) > 4 else v, (" " in v or "\t" in v or "\n" in v or "\r" in v))

    def token(self):
        """tenant_access_token 有效期 2 小时，剩余不足 5 分钟时提前续签。"""
        now = time.time()
        if self._token and now < self._expire_at - 300:
            return self._token
        payload = {
            "app_id": self.app_id,
            "app_secret": self.app_secret,
        }
        r = self._request("POST", "/auth/v3/tenant_access_token/internal", payload)
        if r.get("code") != 0:
            # 诊断：打印两个参数的长度与首尾字符，便于定位是值被污染还是格式不对
            print("凭证诊断 app_id:     %s" % self._fingerprint(self.app_id), file=sys.stderr)
            print("凭证诊断 app_secret: %s" % self._fingerprint(self.app_secret), file=sys.stderr)
            raise RuntimeError("获取 token 失败: code=%s msg=%s" % (r.get("code"), r.get("msg")))
        self._token = r["tenant_access_token"]
        self._expire_at = now + int(r.get("expire", 7200))
        return self._token

    def fetch(self, table_id):
        out = []
        page_token = None
        while True:
            params = {"page_size": "500"}
            if page_token:
                params["page_token"] = page_token
            path = "/bitable/v1/apps/%s/tables/%s/records?%s" % (
                BASE_TOKEN, table_id, urllib.parse.urlencode(params))
            # 取记录是 GET，且必须带上 Authorization 头
            r = self._request("GET", path, token=self.token())
            if r.get("code") != 0:
                raise RuntimeError("读取表 %s 失败: code=%s msg=%s" % (
                    table_id, r.get("code"), r.get("msg")))
            data = r.get("data", {})
            for item in data.get("items", []):
                rec = dict(item.get("fields", {}))
                rec.pop("record_id", None)
                out.append(rec)
            if not data.get("has_more"):
                break
            page_token = data.get("page_token")
            if not page_token:
                break
            time.sleep(0.3)   # 页间加间隔，规避 QPS 限流
        return out


class CliBackend:
    """本地调试：走已登录的 lark-cli。"""

    def __init__(self):
        env_path = os.environ.get("LARK_CLI")
        if env_path:
            self.cli = env_path
        else:
            self.cli = shutil.which("lark-cli")
            if not self.cli and os.name == "nt":
                self.cli = shutil.which("lark-cli.cmd")
            if not self.cli:
                raise RuntimeError("未找到 lark-cli，请设置 LARK_CLI 环境变量")

    def fetch(self, table_id):
        out = []
        offset = 0
        page_size = 200   # CLI 单页上限为 200
        while True:
            cmd = [
                self.cli, "base", "+record-list",
                "--base-token", BASE_TOKEN,
                "--table-id", table_id,
                "--as", "user",
                "--format", "json",
                "--limit", str(page_size),
                "--offset", str(offset),
            ]
            proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
            raw = proc.stdout.strip()
            if not raw:
                raise RuntimeError("lark-cli 无输出: %s" % proc.stderr.strip())
            payload = json.loads(raw)
            if not payload.get("ok"):
                raise RuntimeError("接口失败: %s" % json.dumps(payload.get("error"), ensure_ascii=False))

            data = payload["data"]
            fields = data.get("fields", [])
            for row in data.get("data", []):
                rec = {}
                for name, val in zip(fields, row):
                    if name == "record_id":
                        continue
                    rec[name] = val
                out.append(rec)

            if not data.get("has_more"):
                break
            if not data.get("data"):
                break
            offset += len(data["data"])
        return out


def make_backend():
    app_id = os.environ.get("LARK_APP_ID", "").strip()
    app_secret = os.environ.get("LARK_APP_SECRET", "").strip()
    if app_id and app_secret:
        print("取数模式：OpenAPI（凭证来自环境变量）", file=sys.stderr)
        return OpenApiBackend(app_id, app_secret)
    print("取数模式：lark-cli（本地调试）", file=sys.stderr)
    return CliBackend()


# ---------------------------------------------------------------- 字段取值

def pick_str(v, default=None):
    """文本 / 单选 / 公式：取出干净字符串。"""
    if v is None:
        return default
    if isinstance(v, list):
        if not v:
            return default
        v = v[0]
    if isinstance(v, dict):
        v = v.get("value", v.get("text", default))
    if v is None:
        return default
    s = str(v).strip()
    return s or default


def pick_list(v):
    """多选：字符串数组。"""
    if v is None:
        return []
    if isinstance(v, list):
        out = []
        for x in v:
            if isinstance(x, dict):
                x = x.get("text", x.get("value"))
            if x:
                out.append(str(x).strip())
        return [x for x in out if x]
    return [str(v).strip()]


def pick_date(v):
    """日期：毫秒时间戳或 ISO 字符串 -> date 对象。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return datetime.fromtimestamp(v / 1000, CST)
    if isinstance(v, dict):
        v = v.get("value")
        return pick_date(v)
    s = str(v).strip()
    if not s:
        return None
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        pass
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(s, fmt).replace(tzinfo=CST)
        except ValueError:
            continue
    return None


# ---------------------------------------------------------------- 统计口径

STAGE_DONE, STAGE_DOING, STAGE_TODO, STAGE_HOLD = "完成", "进行中", "未开始", "暂不推进"

# 区域推广进度表的 5 个推进环节（顺序即漏斗顺序）
STAGES = ["文员提需", "HR发单", "供应商+业务培训", "推广结束", "运营数据晾晒"]


def build_progress(rows):
    records = []
    for r in rows:
        states = {s: pick_str(r.get(s), STAGE_TODO) for s in STAGES}
        records.append({
            "code": pick_str(r.get("编码"), "—"),
            "region": pick_str(r.get("区域"), "未分类"),
            "site": pick_str(r.get("站点/HUB"), "—"),
            "states": states,
        })

    total = len(records)

    def stage_count(stage, value):
        return sum(1 for x in records if x["states"].get(stage) == value)

    # 按区域聚合：已完成 HUB = 五个环节全部为「完成」
    by_region = collections.defaultdict(lambda: {"total": 0, "done": 0, "doing": 0, "todo": 0, "hold": 0})
    for x in records:
        b = by_region[x["region"]]
        b["total"] += 1
        vals = list(x["states"].values())
        if all(v == STAGE_DONE for v in vals):
            b["done"] += 1
        elif any(v == STAGE_HOLD for v in vals):
            b["hold"] += 1
        elif any(v == STAGE_DOING for v in vals):
            b["doing"] += 1
        else:
            b["todo"] += 1

    regions = []
    for name, b in by_region.items():
        regions.append({
            "name": name,
            "total": b["total"],
            "done": b["done"],
            "doing": b["doing"],
            "todo": b["todo"],
            "hold": b["hold"],
            "rate": round(b["done"] / b["total"] * 100, 1) if b["total"] else 0.0,
        })
    regions.sort(key=lambda x: (-x["rate"], -x["total"]))

    # 整体完成率 = 各区域完成率的等权平均（每个区域权重相同，不论 HUB 多少）
    overall = round(sum(x["rate"] for x in regions) / len(regions), 1) if regions else 0.0

    # 参考口径：按 HUB 数加权的完成率，用于对照说明
    weighted = round(sum(x["done"] for x in regions) / total * 100, 1) if total else 0.0

    # 漏斗：每个环节的完成数
    funnel = [{"name": s, "value": stage_count(s, STAGE_DONE)} for s in STAGES]

    # 每个环节的状态构成
    stage_matrix = []
    for s in STAGES:
        stage_matrix.append({
            "name": s,
            "完成": stage_count(s, STAGE_DONE),
            "进行中": stage_count(s, STAGE_DOING),
            "未开始": stage_count(s, STAGE_TODO),
            "暂不推进": stage_count(s, STAGE_HOLD),
        })

    return {
        "total_sites": total,
        "region_count": len(by_region),
        "overall_rate": overall,      # 区域等权平均（看板主口径）
        "weighted_rate": weighted,    # HUB 加权（对照口径）
        "stage_done": {s: stage_count(s, STAGE_DONE) for s in STAGES},
        "regions": regions,
        "funnel": funnel,
        "stage_matrix": stage_matrix,
        "records": sorted(records, key=lambda x: (x["region"], x["site"])),
    }


def build_plan(rows):
    total = len(rows)

    def dist(field, unwrap=True):
        c = collections.Counter()
        for r in rows:
            v = r.get(field)
            if unwrap:
                v = pick_str(v, "未填写")
                c[v] += 1
            else:
                for x in pick_list(v) or ["未填写"]:
                    c[x] += 1
        return c

    status = dist("推广状态")
    region_raw = dist("区域")

    # 归一化区域名：去掉尾部的编号括号，如「德州大区（8）」->「德州大区」
    region_c = collections.Counter()
    for k, v in region_raw.items():
        name = k.split("（")[0].split("(")[0].strip() or "未分类"
        region_c[name] += v

    suppliers = [k for k in dist("供应商").keys() if k != "未填写"]
    sites = set(pick_str(r.get("站点/HUB")) for r in rows if pick_str(r.get("站点/HUB")))

    # 区域 x 状态 交叉
    cross = collections.defaultdict(collections.Counter)
    for r in rows:
        reg = (pick_str(r.get("区域"), "未分类")).split("（")[0].split("(")[0].strip() or "未分类"
        cross[reg][pick_str(r.get("推广状态"), "未填写")] += 1

    batch = collections.Counter()
    for r in rows:
        for b in pick_list(r.get("推广批次")) or ["未填写"]:
            batch[b] += 1

    return {
        "total_tasks": total,
        "site_count": len(sites),
        "supplier_count": len(suppliers),
        "status_dist": dict(status),
        "region_dist": dict(region_c),
        "region_status_matrix": {k: dict(v) for k, v in cross.items()},
        "batch_dist": dict(batch),
    }


# ---------------------------------------------------------------- 主流程

def main():
    probe = "--probe" in sys.argv
    backend = make_backend()

    print("[1/3] 拉取「2、区域推广进度」…", file=sys.stderr)
    prog_rows = backend.fetch(TABLE_PROGRESS)
    print("      %d 条" % len(prog_rows), file=sys.stderr)
    time.sleep(0.5)   # 表间加间隔，规避 QPS 限流

    print("[2/3] 拉取「1、推广计划」…", file=sys.stderr)
    plan_rows = backend.fetch(TABLE_PLAN)
    print("      %d 条" % len(plan_rows), file=sys.stderr)

    print("[3/3] 聚合…", file=sys.stderr)
    prog = build_progress(prog_rows)
    plan = build_plan(plan_rows)

    now = datetime.now(CST)
    payload = {
        "updated_at": now.isoformat(timespec="seconds"),
        "generated_ts": int(now.timestamp()),
        "source": {
            "base_token": BASE_TOKEN,
            "tables": [
                {"id": TABLE_PROGRESS, "name": "2、区域推广进度", "rows": len(prog_rows)},
                {"id": TABLE_PLAN, "name": "1、推广计划", "rows": len(plan_rows)},
            ],
        },
        "kpi": {
            "progress": {
                "total_sites": prog["total_sites"],
                "region_count": prog["region_count"],
                "overall_rate": prog["overall_rate"],
                "weighted_rate": prog["weighted_rate"],
                "stage_done": prog["stage_done"],
            },
            "plan": {
                "total_tasks": plan["total_tasks"],
                "site_count": plan["site_count"],
                "supplier_count": plan["supplier_count"],
            },
        },
        "progress": prog,
        "plan": plan,
    }

    if probe:
        print(json.dumps(payload["kpi"], ensure_ascii=False, indent=2))
        return

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, default=str)
    print("已生成 %s（%s）" % (OUT_PATH, payload["updated_at"]), file=sys.stderr)


if __name__ == "__main__":
    main()
