"""
asin_variant_score.py — ASIN 变体流量得分查询
输入：data/asin_variant_score/input/asin_list.txt（每行一个 ASIN）
输出：data/asin_variant_score/output/asin_variant_score.xlsx

流程：
  1. 读取并校验 ASIN 列表
  2. 逐个查询变体，按 parentAsin 去重（兄弟变体只记录一次）
  3. 分批（≤100）查询所有子 ASIN 的流量得分
  4. 按参考格式写入 Excel（D列=ASIN，G列=得分，含公式）
"""

import hashlib
import json
import os
import time
from pathlib import Path

import pandas as pd
import requests

# ── 西柚 API 配置 ─────────────────────────────────────────
XIYOU_BASE_URL     = "https://openapi.xiyouzhaoci.com"
XIYOU_CLIENT_ID    = "xiyou.ak.8273645"
XIYOU_CLIENT_SECRET = "xiyou.sk.928374650192837"
COUNTRY = "US"

# 流量得分接口每批上限
TRAFFIC_BATCH_SIZE = 100
# 变体查询间隔（秒），避免触发限流
VARIANT_INTERVAL   = 0.3
# 请求失败重试次数
MAX_RETRIES = 3

# ── 输出 Excel 格式 ───────────────────────────────────────
HEADERS = ['状态', '广告活动', '广告组', 'ASIN', '存在', '', '西柚找词得分',
           '热度', '相关性', '竞价策略', '品牌', '备注', '批量插入']

COLUMN_WIDTHS = {
    'A': 8,  'B': 30, 'C': 15, 'D': 15, 'E': 8,
    'F': 5,  'G': 12, 'H': 8,  'I': 8,  'J': 10,
    'K': 8,  'L': 8,  'M': 25,
}


# ── 公共入口 ──────────────────────────────────────────────

def run():
    base_path  = Path(__file__).resolve().parents[1]
    input_dir  = base_path / 'data' / 'asin_variant_score' / 'input'
    output_dir = base_path / 'data' / 'asin_variant_score' / 'output'
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    # 查找 txt 输入文件（优先 asin_list.txt，否则取第一个 .txt）
    txt_files = list(input_dir.glob('*.txt'))
    if not txt_files:
        print(f"❌ 未找到输入文件，请上传 .txt 文件到：{input_dir}")
        return
    input_file = next((f for f in txt_files if f.name == 'asin_list.txt'), txt_files[0])
    print(f"📄 输入文件：{input_file.name}")

    # 1. 读取并校验 ASIN 列表
    asins = _load_asins(input_file)
    if not asins:
        print("❌ 输入文件为空")
        return
    print(f"📌 读取 {len(asins)} 个 ASIN")

    # 2. 获取变体，按 parentAsin 去重
    variant_groups = _fetch_all_variants(asins)
    if not variant_groups:
        print("❌ 未获取到任何变体数据")
        return
    print(f"📌 去重后共 {len(variant_groups)} 个父ASIN组")

    # 3. 按父体分组批量查询流量得分
    total_children = sum(len(g['childAsins']) for g in variant_groups)
    print(f"📌 共 {total_children} 个子ASIN，按父体批量查询得分...")

    score_map = _fetch_scores_by_group(variant_groups)
    print(f"📌 得分查询完成，获取 {len(score_map)} 条记录")

    # 5. 写入 Excel
    output_path = output_dir / 'asin_variant_score.xlsx'
    _write_excel(variant_groups, score_map, output_path)
    print(f"✅ 已保存：{output_path.name}")


# ── API 调用 ──────────────────────────────────────────────

def _api_call(endpoint, body, retry=MAX_RETRIES):
    """带签名的 POST 请求，支持自动重试"""
    for attempt in range(1, retry + 1):
        try:
            timestamp = str(int(time.time()))
            body_str  = json.dumps(body, separators=(',', ':'), sort_keys=True)
            raw  = f"{XIYOU_CLIENT_ID}{timestamp}{XIYOU_CLIENT_SECRET}{body_str}"
            sign = hashlib.sha256(raw.encode('utf-8')).hexdigest()

            headers = {
                'Content-Type': 'application/json',
                'X-Client-Id':  XIYOU_CLIENT_ID,
                'X-Timestamp':  timestamp,
                'X-Sign':       sign,
            }
            resp = requests.post(
                f"{XIYOU_BASE_URL}{endpoint}",
                headers=headers, json=body, timeout=30
            )
            resp.raise_for_status()
            return resp.json()
        except Exception as e:
            if attempt < retry:
                wait = attempt * 1.5
                print(f"  ⚠ 第{attempt}次请求失败，{wait:.0f}s 后重试：{e}")
                time.sleep(wait)
            else:
                raise


def _fetch_variant(asin):
    """查询单个 ASIN 的变体信息"""
    try:
        return _api_call("/v1/asins/variations", {"country": COUNTRY, "asin": asin})
    except Exception as e:
        print(f"  ❌ 变体接口失败 ASIN={asin}: {e}")
        return None


def _fetch_scores_by_group(variant_groups):
    """
    按父体分组批量查询流量得分。
    同一父体的所有子ASIN一次性发送（超过100个时自动拆批）。
    返回 { asin: totalTrafficScore }
    """
    score_map = {}

    for i, group in enumerate(variant_groups, 1):
        parent   = group['parentAsin']
        children = group['childAsins']
        if not children:
            continue

        print(f"  → [{i}/{len(variant_groups)}] 父ASIN {parent}，查询 {len(children)} 个子ASIN得分")

        # 子ASIN超过100时拆批（通常不会触发）
        batches = [children[j:j + TRAFFIC_BATCH_SIZE]
                   for j in range(0, len(children), TRAFFIC_BATCH_SIZE)]

        for batch in batches:
            try:
                body = {"entities": [{"country": COUNTRY, "asin": a} for a in batch]}
                data = _api_call("/v1/asins/traffic", body)
                for item in data.get("entities", []):
                    score_map[item["asin"]] = item.get("totalTrafficScore", 0)
            except Exception as e:
                print(f"  ❌ 父ASIN {parent} 得分查询失败: {e}")

    return score_map


# ── 业务逻辑 ──────────────────────────────────────────────

def _load_asins(filepath):
    """读取并去重 ASIN 列表，过滤格式不合法的行"""
    seen, result = set(), []
    with open(filepath, 'r', encoding='utf-8') as f:
        for line in f:
            asin = line.strip().upper()
            if not asin:
                continue
            if len(asin) != 10:
                print(f"  ⚠ 跳过格式异常（非10位）: {asin}")
                continue
            if asin in seen:
                continue
            seen.add(asin)
            result.append(asin)
    return result


def _fetch_all_variants(asins):
    """
    逐个查询变体，按 parentAsin 去重。
    兄弟变体（相同 parentAsin）只保留第一个查询到的组。
    """
    seen_parents = set()
    groups = []
    total  = len(asins)

    for i, asin in enumerate(asins, 1):
        print(f"  → [{i}/{total}] 查询变体: {asin}")
        data = _fetch_variant(asin)

        if data is None:
            # 接口失败：将该 ASIN 单独作为一组保留
            if asin not in seen_parents:
                seen_parents.add(asin)
                groups.append({"parentAsin": asin, "childAsins": [asin]})
            continue

        parent      = data.get("parentAsin") or asin
        child_asins = list(data.get("childAsins") or [])

        # 单品无变体：把自身加入子列表
        if not data.get("parentAsin") and asin not in child_asins:
            child_asins.append(asin)

        if parent in seen_parents:
            print(f"    ↳ 已有父ASIN {parent}，跳过（兄弟变体去重）")
        else:
            seen_parents.add(parent)
            groups.append({"parentAsin": parent, "childAsins": child_asins})

        # 变体查询间隔，避免触发限流
        if i < total:
            time.sleep(VARIANT_INTERVAL)

    return groups


# ── Excel 输出 ────────────────────────────────────────────

def _write_excel(variant_groups, score_map, output_path):
    rows     = []
    excel_row = 2   # 第1行为表头，数据从第2行起

    for group in variant_groups:
        children     = group["childAsins"]
        num_variants = len(children)

        # 按得分降序排列
        children_sorted = sorted(children,
                                 key=lambda a: score_map.get(a, 0),
                                 reverse=True)
        total_score = sum(score_map.get(a, 0) for a in children_sorted)

        # ── 汇总行（广告组行）：D 列留空，G 列为组总分 ──
        r = excel_row
        rows.append([
            '',                                                                         # A 状态
            f'=IF(AND(C{r}<>"",H{r}<>""),CONCAT("DL_",C{r},"_精准_",H{r}),"")' ,     # B 广告活动
            f'{children_sorted[0]}_{num_variants}V',                                   # C 广告组
            '',                                                                         # D ASIN（汇总行留空）
            'Y',                                                                        # E 存在
            '',                                                                         # F 空列
            total_score,                                                                # G 西柚找词得分（组总分）
            f'=IF(D{r}="",IF(G{r}="","",IF(G{r}>2500,"VH",IF(G{r}>1500,"H",'
            f'IF(G{r}>500,"M",IF(G{r}>250,"L",IF(G{r}>=0,"VL","")))))),"")',           # H 热度
            '',                                                                         # I 相关性
            '',                                                                         # J 竞价策略
            '',                                                                         # K 品牌
            '',                                                                         # L 备注
            '',                                                                         # M 批量插入
        ])
        excel_row += 1

        # ── 子 ASIN 行：D 列=ASIN，G 列=个体得分 ──
        for child in children_sorted:
            r = excel_row
            rows.append([
                '',                                                                     # A 状态
                '',                                                                     # B 广告活动
                '',                                                                     # C 广告组
                child,                                                                  # D ASIN
                'Y',                                                                    # E 存在
                '',                                                                     # F 空列
                score_map.get(child, 0),                                                # G 西柚找词得分
                '',                                                                     # H 热度
                '',                                                                     # I 相关性
                '',                                                                     # J 竞价策略
                '',                                                                     # K 品牌
                '',                                                                     # L 备注
                f'=IF(D{r}="","",CONCATENATE("asin=","""",D{r},""""))',                 # M 批量插入
            ])
            excel_row += 1

    df = pd.DataFrame(rows, columns=HEADERS)

    if os.path.exists(output_path):
        os.remove(output_path)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        ws = writer.sheets['Sheet1']
        for col, width in COLUMN_WIDTHS.items():
            ws.column_dimensions[col].width = width

    print(f"  写入 {len(rows)} 行，{len(variant_groups)} 个父ASIN组")
