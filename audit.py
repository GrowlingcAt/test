# -*- coding: utf-8 -*-
"""按 README.docx 规则审计 答案.json / 答案.docx 是否合规"""
import json, re, os, docx

BASE = "/Users/xxx/code/test1"
CANNOT = "该需求不能直接转换为FOFA搜索语句"
PKG = "pkg-08e82c8d"

# ---------- 1. 重读 README 规则原文 ----------
rd = docx.Document(f"{BASE}/README.docx")
rules = [p.text for p in rd.paragraphs if p.text.strip()]
key_rules = [t for t in rules if any(k in t for k in
             ["固定答案", "完全一致", "不得缺题", "UTF-8", "8 MB", "不得包含", "JSONL", "建议命名"])]
print("=" * 30, "README 关键规则原文", "=" * 30)
for t in key_rules:
    print(" -", t)

# ---------- 2. 读入题目与答案 ----------
def ids_of(path, key="题号"):
    out = []
    for p in docx.Document(path).paragraphs:
        m = re.search(rf'"{key}":\s*"([^"]+)"', p.text)
        if m:
            out.append(m.group(1))
    return out

q_ids = ids_of(f"{BASE}/题目.docx")
raw = open(f"{BASE}/答案.json", "rb").read()
data = json.loads(raw.decode("utf-8"))
ans = data["答案"]
a_ids = [a["题号"] for a in ans]

print("\n" + "=" * 30, "审计结果", "=" * 30)
def chk(ok, name, extra=""):
    print(f"{'✅' if ok else '❌'} {name}{(' — ' + extra) if extra else ''}")
    return ok

allok = True
# 顶层字段
allok &= chk(set(data.keys()) == {"选手名称", "参赛包编号", "答案"}, "顶层字段为 选手名称/参赛包编号/答案", str(list(data.keys())))
allok &= chk(data.get("参赛包编号") == PKG, f"参赛包编号保持不变({PKG})", data.get("参赛包编号"))
allok &= chk(isinstance(ans, list), "答案 是 JSON 数组(非JSONL)")

# 题号完整性
allok &= chk(len(ans) == len(q_ids) == 100, f"题号数量一致(答案{len(ans)}/题目{len(q_ids)})")
allok &= chk(a_ids == q_ids, "题号顺序与题目完全一致")
allok &= chk(len(set(a_ids)) == len(a_ids), "无重复题号")
allok &= chk(set(a_ids) == set(q_ids), "无缺题、无未知题号")

# 每条结构
bad_struct = [a for a in ans if set(a.keys()) != {"题号", "查询语句"}]
allok &= chk(not bad_struct, "每条仅含 题号/查询语句 两个字段", str(bad_struct[:2]))

# 查询语句内容
empty = [a["题号"] for a in ans if not str(a["查询语句"]).strip()]
allok &= chk(not empty, "无空值查询语句", str(empty))

multi_line = [a["题号"] for a in ans if "\n" in str(a["查询语句"])]
allok &= chk(not multi_line, "查询语句均为单行(无换行)", str(multi_line))

markdown = [a["题号"] for a in ans if any(x in str(a["查询语句"]) for x in ["```", "**", "##"])]
allok &= chk(not markdown, "无 Markdown 代码块/标记", str(markdown))

# 非固定答案的必须是单条 FOFA 查询(不能夹带解释性中文句子)
def looks_like_query(s):
    # 固定答案放行
    if s == CANNOT:
        return True
    # 允许中文出现在引号内的检索值里，但不能出现引导性解释词
    for w in ["应该", "可以先用", "建议", "解释", "即：", "注意：", "因为"]:
        if w in s:
            return False
    # 必须以字段= 形式开头(允许括号包裹的整体)
    return bool(re.match(r'^[\s(]*(?:[a-zA-Z_][\w.]*)\s*(?:=|==|!=|>|<|\*=)', s))

bad_fmt = [a["题号"] for a in ans if not looks_like_query(a["查询语句"])]
allok &= chk(not bad_fmt, "每条均为合法 FOFA 字段语法或固定答案", str(bad_fmt))

cannot_ids = [a["题号"] for a in ans if a["查询语句"] == CANNOT]
ok_ids = [a["题号"] for a in ans if a["查询语句"] != CANNOT]
print(f"   固定答案 {len(cannot_ids)} 条: {cannot_ids}")
print(f"   FOFA 查询 {len(ok_ids)} 条")

# 编码与体积
allok &= chk(len(raw) < 8 * 1024 * 1024, "文件 < 8MB", f"{len(raw)/1024:.1f} KB")
allok &= chk(raw[:3] != b"\xef\xbb\xbf" or True, "UTF-8 编码(BOM可选)", "utf-8")

# 敏感信息
secret = [a["题号"] for a in ans if re.search(r'(api[_-]?key|secret|password|token\s*=)', str(a["查询语句"]), re.I)]
allok &= chk(not secret, "无账号/密钥/未授权数据", str(secret))

# ---------- 3. 审计 docx 输出 ----------
print("\n" + "=" * 28, "答案.docx 结构审计", "=" * 28)
d = docx.Document(f"{BASE}/答案.docx")
txt = "\n".join(p.text for p in d.paragraphs)
allok &= chk(PKG in txt, f"docx 内保留参赛包编号({PKG})")
filled = len(re.findall(r'"查询语句":\s*"[^"]+"', txt))
allok &= chk(filled == 100, "docx 内 100 条查询语句均已填充(非空)", f"{filled}/100")
blank = re.findall(r'"查询语句":\s*""', txt)
allok &= chk(not blank, "docx 内无留空的查询语句", str(len(blank)))

print("\n" + ("🎉 全部通过：生成物符合 README.docx 全部硬性要求" if allok else "⚠️ 存在未通过项，见上"))
