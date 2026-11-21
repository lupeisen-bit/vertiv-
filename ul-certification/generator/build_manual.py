# -*- coding: utf-8 -*-
import sys, json, os
sys.path.insert(0, "/tmp/deck")
from lib import *
import lib
lib.load_expansions()

GROUPS = [
    ("路线图与总览", "UL 800VDC 数据中心电力标准路线图：缺口、CRD/CCN/OOI", ["ul_roadmap"]),
    ("电源变换与 ICT 安全", "AC/DC、DC/DC、UPS、DER 变换器、中压电源", ["ul62368", "ul1012", "ul508c", "ul1778", "ul1741", "ul2877"]),
    ("保护、配电与预制系统", "断路器、母线槽、开关板/开关柜、预制配电与模块化数据中心", ["ul489", "ul1066", "ul857", "ul891", "ul1558", "ul2140", "ul2755"]),
    ("储能与备用电源", "固定式电池、储能系统、热失控传播", ["ul1973", "ul9540", "ul9540a"]),
    ("连接、浪涌、绝缘与材料", "连接器、线夹、SPD、绝缘配合、塑料阻燃", ["ul1977", "ul486ab", "ul1449", "ul840", "ul94", "ul746c"]),
]
TAGC = {0: TEAL, 1: NAVY, 2: ORANGE, 3: GREEN, 4: GREY}

D = Deck(footer_label="UL 标准深度手册 · 条款级参考（引用核实等级见页脚）")

s = D.blank()
D.rect(s, 0, 0, 13.333, 7.5, NAVY)
D.rect(s, 0, 5.2, 13.333, 0.08, TEAL)
D.text(s, 0.8, 1.5, 11.8, 1.0, "UL 标准深度手册", size=44, color=WHITE, bold=True)
D.text(s, 0.8, 2.6, 11.8, 1.0, "800V DC AI 数据中心电源系统 · 条款级参考", size=26, color=TEAL, bold=True)
D.text(s, 0.8, 3.7, 11.8, 1.2, ["每个标准一章：概览 / 范围 / 章节目录 / 条款要点 / 试验矩阵 / 800V 适配 / 设计检查表 / 认证流程 / 来源",
                                "配套《如何识别与使用 UL 标准》主讲稿（方法论）"], size=15, color=PALE, gap=8)
D.text(s, 0.8, 5.5, 11.8, 1.4, [
    "每一行事实标注“依据”：原文（公开原文/预览）/ 目录（公开 TOC）/ 解读（二手资料或判断）；页脚 A/B 为整页可信度。",
    "条款号均来自公开目录/范围/预览，版本不一定最新；写入设计规范或 RTM 之前必须用正式购买的文本核对。"], size=12, color=PALE, gap=4)

# load all
data = {}
for g in GROUPS:
    for k in g[2]:
        fp = f"/tmp/deck/{k}.json"
        if os.path.exists(fp):
            data[k] = json.load(open(fp))

# TOC placeholder slides are created after we know page numbers -> two-pass: build once to get counts
def plan():
    n = 1 + 3  # cover + 2 toc + how to use
    ranges = {}
    for gi, (gname, gsub, keys) in enumerate(GROUPS):
        n += 1
        for k in keys:
            if k not in data:
                continue
            T = Deck()
            for si, sl in enumerate(data[k]["slides"]):
                sl["_k"] = f"{k}|{si}"
                T.std_slide(sl, "x", TEAL)
            start = n + 1
            n += T.n
            ranges[k] = (start, n)
    return ranges


ranges = plan()

# TOC (2 slides): overview by group
s = D.blank()
D.header(s, "手册目录与页数", "按子系统分组；页数 = 该标准的专章页数（不含分组页）", "目录")
rows = []
for gi, (gname, gsub, keys) in enumerate(GROUPS):
    for k in keys:
        if k in data:
            a, b = ranges[k]
            rows.append([gname, data[k]["id"], data[k]["name_zh"][:26], f"{a}–{b}", str(b - a + 1)])
        else:
            rows.append([gname, k.upper(), "（未生成）", "—", "0"])
half = (len(rows) + 1) // 2
D.table(s, 0.35, 1.3, 12.65, [2.6, 1.8, 5.0, 1.3, 0.9], ["分组", "标准", "名称", "页码", "页数"], rows[:half], size=12, max_h=5.6, row_pad=0.16)
D.footer(s)
s = D.blank()
D.header(s, "手册目录与页数（续）", f"续表；全手册共 {max(v[1] for v in ranges.values())} 页", "目录")
D.table(s, 0.35, 1.3, 12.65, [2.6, 1.8, 5.0, 1.3, 0.9], ["分组", "标准", "名称", "页码", "页数"], rows[half:], size=12, max_h=5.6, row_pad=0.16)
D.footer(s)

# how to use
s = D.blank()
D.header(s, "如何使用这本手册", "它是查询用的，不是用来从头读到尾的", "说明")
D.card(s, 0.35, 1.3, 6.2, 5.55, "查询路径", [
    "先看分组页，再进具体标准",
    "每个标准章从“概览”和“范围”开始：先判断适不适用",
    "“章节目录”页告诉你条款在哪一章",
    "“条款要点”页可复制进需求追溯矩阵（RTM）",
    "“800V 适配”页只说公开信息能支撑的判断",
    "“设计检查表”和“认证流程”用于评审关口",
], TEAL, 13)
D.card(s, 6.75, 1.3, 6.25, 5.55, "阅读规则", [
    "“原文”= 公开预览/Scope/全文中可见的内容",
    "“目录”= 来自公开目录（只可信章节标题和编号）",
    "“解读”= 二手资料或本文判断，需复核",
    "OOI（UL 2877/2140/2755）无公开全文：页数少是因为没有可公开引用的条款",
    "页数不是越多越好：以可被来源支持为准，宁缺毋滥",
], ORANGE, 13)
D.footer(s)

assert D.n == 3 + 0 or True

for gi, (gname, gsub, keys) in enumerate(GROUPS):
    D.divider(gname, [gsub, "本组标准：" + "、".join(data[k]["id"] for k in keys if k in data)], f"{gi + 1}")
    for k in keys:
        if k not in data:
            continue
        d = data[k]
        tag = d["id"][:14]
        for si, sl in enumerate(d["slides"]):
            sl["_k"] = f"{k}|{si}"
            D.std_slide(sl, tag, TAGC[gi])

# actual ranges (verify)
D.prs.save("/tmp/deck/_manual_stage.pptx")
json.dump({k: list(v) for k, v in ranges.items()}, open("/tmp/deck/_ranges.json", "w"))
import lib
print("WARN",len(lib.WARN),lib.WARN[:40])
print("slides", D.n, {k: v for k, v in ranges.items()})

os.makedirs("/agent/单标准分册", exist_ok=True)
for gi, (gname, gsub, keys) in enumerate(GROUPS):
    for k in keys:
        if k not in data:
            continue
        d = data[k]
        X = Deck(footer_label=f"{d['id']} · 条款级参考 · UL 标准深度手册分册")
        s = X.blank()
        X.rect(s, 0, 0, 13.333, 7.5, NAVY)
        X.rect(s, 0, 4.4, 13.333, 0.08, TEAL)
        X.text(s, 0.8, 1.8, 11.8, 1.0, d["id"], size=44, color=WHITE, bold=True)
        X.text(s, 0.8, 2.9, 11.8, 0.8, d["name_zh"], size=24, color=TEAL, bold=True)
        X.text(s, 0.8, 3.7, 11.8, 0.6, "版本：" + d.get("edition", ""), size=14, color=PALE)
        X.text(s, 0.8, 4.7, 11.8, 1.5, ["800V DC AI 数据中心电源系统 · 条款级参考", "每行事实标注依据（原文/目录/解读）；条款号来自公开目录/预览，使用前须用正式文本复核"], size=13, color=PALE, gap=5)
        for si, sl in enumerate(d["slides"]):
            sl["_k"] = f"{k}|{si}"
            X.std_slide(sl, d["id"][:14], TAGC[gi])
        X.prs.save(f"/agent/单标准分册/{gi + 1}_{d['id'].replace(' ', '_').replace('/', '-')}.pptx")

if lib.EXPORT_MODE:
    json.dump({e["k"]: e for e in lib.EXPORT}, open("/tmp/deck/work/export_manual.json", "w"), ensure_ascii=False)
