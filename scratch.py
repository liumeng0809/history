import re
import time
from openai import OpenAI
import httpx
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# 解析 **加粗** 行内标记；非贪婪，避免跨段误吞
_BOLD_RE = re.compile(r'\*\*(.+?)\*\*')


def add_runs_with_bold(paragraph, text):
    """把含 **加粗** 的文本拆成多个 run，加粗部分 bold=True。"""
    # 对有 ** 成对出现的段落，保留加粗；对未配对的 ** 标记直接清除
    if text.count('**') >= 2:
        # 正常解析成对 **
        pos = 0
        for m in _BOLD_RE.finditer(text):
            if m.start() > pos:
                paragraph.add_run(text[pos:m.start()])
            run = paragraph.add_run(m.group(1))
            run.bold = True
            pos = m.end()
        if pos < len(text):
            paragraph.add_run(text[pos:])
    else:
        paragraph.add_run(text.replace('**', ''))

# ================= 配置区域 =================
# 把这里替换成你刚才在硅基流动复制的 API 密钥
API_KEY = "sk-Qaes2KCUDHr67GZoif13ySqsuFsD7NYWAnLoVcNcAlv3mcXW"
# 大模型名称（硅基流动的免费强模型）
MODEL_NAME = "DeepSeek-V4-Flash"
# 你想生成的科普主题
TOPIC = "东晋历史"
# ============================================

# 初始化大模型客户端
client = OpenAI(
    api_key=API_KEY,
    base_url="http://aiserver.hisi.huawei.com/v1",
    # 关键修复：绕过 Windows 注册表里的 IE 系统代理（华为 proxycn2:8080），
    # 否则 httpx 会走代理 → 504 Gateway Time-out。codeAi 的配置不走代理可用。
    http_client=httpx.Client(trust_env=False)
)


def call_ai(prompt):
    """调用大模型的通用函数"""
    print("正在思考中...", end="")
    response = client.chat.completions.create(
        model=MODEL_NAME,
        messages=[{"role": "user", "content": prompt}]
    )
    print("完成!")
    return response.choices[0].message.content


def generate_outline(topic):
    """第一步：生成详细大纲"""
    prompt = f"""
    你是一位严谨的历史学家。请为主题【{topic}】生成一份极其详细的时间线大纲。
    要求：
    1. 从建立到灭亡，按时间顺序排列。
    2. 必须细化到每一个皇帝、关键权臣、重大战役、重大政治事件。
    3. 只输出大纲列表，不要写正文内容。每个大纲条目单独占一行。
    """
    print(f"\n第一阶段：生成【{topic}】的详细大纲...")
    return call_ai(prompt)


def generate_section_details(outline_line, prev_detail):
    """第二步：针对大纲的每一行，生成详细正文。
    prev_detail: 上一条已生成正文的尾部，用于自然衔接（第一条为空字符串）。"""
    prev_block = ""
    if prev_detail:
        prev_block = f"""

【上一条大纲的正文结尾】（这是上一条讲到的地方，仅供你知道故事进行到哪了）：
{prev_detail}

【衔接要求——重要】
你的本条要从上面那个地方之后接着讲。只做简短的时间或场景过渡（如"第二年""就在这时""同一年的冬天""王敦的军队还没到建康时"），一句话带过，直接进入本条大纲节点的新内容。
严禁前情回顾：不要复述上一条已经讲过的事件经过、起因或结果；不要用"上回说到""刚才讲到""前文提到""我们之前说过"之类的话。上一条讲过的部分，默认读者已经知道，不重讲。
"""
    prompt = f"""
请针对这一历史节点：【{outline_line}】，写一篇非常详细的科普文章，面向完全不懂历史的普通读者。
要求：
1. 字数不少于 800 字。
2. 聚焦本条大纲节点真正发生的新事，必须含具体年份、人物、经过和影响；
   起因和背景若上一条已交代，本条一句话带过或不提，不要把上一条讲过的内容当起因再重讲一遍。
3. 拒绝宏观叙事，不要泛泛而谈，只要具体史实细节。
4. 直接输出正文内容，使用 Markdown 格式排版（用 ### 作为标题）。
5. 语言口语化、娓娓道来，像跟朋友讲故事，不要学术腔、不要书面语、不要生硬罗列。
6. 与上一条的衔接只做简短的时间/场景过渡，直接进入新内容，严禁前情回顾和复述（详见下方【衔接要求】）。
7. 讲清本条节点的来龙去脉：它发生在大背景下、为什么会发生、牵涉谁、各方什么关系、
   前因后果怎么串——让不懂历史的人听懂这一节点的发展脉络和影响。
   注意：只讲本条节点的新脉络；跨多条大纲的整个事件，前面条目已经讲过的部分不重复。
8. 第一次提到的人名、官职、地名、制度时，顺带用大白话解释一句是什么、干什么的、在哪儿，
   不要默认读者知道（例如'刺史'是地方军政长官、'建康'就是今天的南京）。
9. 【最高原则——严格忠于史实，绝不揣测编造】整篇内容必须基于真实历史，只能写有史料依据的事实。
   不得虚构任何内容，包括但不限于：人物对话、心理活动、细节动作、场景描写、人物结局、事件因果。
   对话不要凭空编写（不要写"某某说：……"除非确有史料记载原话）；人物心理不要揣测（不要写"他心里一沉""他暗想"）；
   场景细节不要脑补（不要写"风很大""烛火摇曳"等无依据描写）。
   拿不准的细节宁可略去或用"据记载""据史书"等留白表述，绝不臆造。
   例如某人是被杀、自杀、逃亡还是病故，必须区分清楚，不能图省事一律写成"被杀"。
10. 第一次提到次要人物时，只交代其确凿的身份与立场（如'刘隗，是晋元帝用来牵制王氏的亲信'），
    不展开未经核对其后人生结局；若确知其结局，简述即可，不添油加醋。
{prev_block}
"""
    print(f"  -> 正在详细扩写: {outline_line}")
    return call_ai(prompt)


def get_tail(text, n=400):
    """取正文结尾 n 字，作为下一条衔接的上下文。"""
    text = text.strip()
    return text[-n:] if len(text) > n else text


def group_outline_by_event(outline_lines):
    """让 AI 把大纲按'完整历史事件'分组，返回 [[行号1,...], [行号2,...], ...]。
    行号从 1 开始，对应 outline_lines 的索引+1。"""
    numbered = "\n".join(f"{i + 1}. {line}" for i, line in enumerate(outline_lines))
    prompt = f"""
下面是一份历史时间线大纲，共 {len(outline_lines)} 条。请按"完整历史事件"把它们分组：
同一个事件的起因、经过、结果要归到同一组，不能把一个事件拆到两个组。
只输出分组结果，格式为每组一行，行号用逗号分隔，例如：
1,2,3
4,5
6,7,8,9
不要输出任何解释、标题或多余文字。

大纲：
{numbered}
"""
    print("\n第一步（补充）：按完整事件给大纲分组...")
    raw = call_ai(prompt)
    groups = []
    for ln in raw.splitlines():
        ln = ln.strip()
        if not ln or not re.match(r'^[\d，,\s]+$', ln):
            continue
        nums = [int(x.strip()) for x in re.split(r'[，,]', ln) if x.strip().isdigit()]
        if nums:
            groups.append(nums)
    # 校验：行号需覆盖且不越界、不重复
    valid, seen = [], set()
    for g in groups:
        clean = [n for n in g if 1 <= n <= len(outline_lines) and n not in seen]
        if clean:
            for n in clean:
                seen.add(n)
            valid.append(clean)
    return valid


def merge_small_groups(groups, max_size=12):
    """贪心合并相邻事件组：当前组若还 ≤ max_size，就把下一组并进来；并后超过 max_size 就不并，当前组独立成文档。
    即"只要还没满 12 条，就继续往里加下一组"。
    groups: [[行号,...], ...]；返回合并后的组列表。"""
    if not groups:
        return groups
    merged = [list(groups[0])]
    for cur in groups[1:]:
        prev = merged[-1]
        if len(prev) + len(cur) <= max_size:
            prev.extend(cur)  # 还没满 12 条，继续并
        else:
            merged.append(list(cur))  # 并了会超 12，当前组独立成文档
    return merged


def markdown_to_word(md_content, topic):
    """第三步：合并并转为 Word 文档"""
    print("\n第三阶段：正在生成 Word 文档...")

    doc = Document()

    # 设置正文中文字体（默认字体不含中文会显示异常）
    style = doc.styles['Normal']
    style.font.name = 'Microsoft YaHei'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
    style.font.size = Pt(12)

    # 大标题（用 Heading 1 样式 + 居中对齐 + 加大字号）
    title = doc.add_heading(f"{topic}详细科普", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title.runs:
        run.font.size = Pt(22)

    # 按 Markdown 的分隔线切分章节，逐块解析
    sections = md_content.split('\n---\n')
    for sec in sections:
        for line in sec.splitlines():
            line = line.rstrip()
            if not line:
                continue
            if line.startswith('### '):
                doc.add_heading(line[4:].strip(), level=2)
            elif line.startswith('## '):
                doc.add_heading(line[3:].strip(), level=1)
            elif line.startswith('# '):
                doc.add_heading(line[2:].strip(), level=0)
            else:
                para = doc.add_paragraph()
                add_runs_with_bold(para, line)
        if sec.strip():
            doc.add_paragraph()  # 章节间空行

    word_file = f"{topic}_详细科普.docx"
    doc.save(word_file)
    print(f"  🎉 已生成: {word_file}")


if __name__ == "__main__":
    # 1. 获取大纲
    outline = generate_outline(TOPIC)

    # 将大纲按行分割成列表，并过滤掉空行
    outline_lines = [line.strip() for line in outline.split('\n') if line.strip()]

    # 清洗大纲：去掉行首序号符号，过滤空行
    clean_lines = []
    for line in outline_lines:
        clean_line = line.lstrip('0123456789.-* ')
        if clean_line:
            clean_lines.append(clean_line)

    total = len(clean_lines)
    print(f"\n大纲共 {total} 条，先全部打印出来：")
    print("=" * 60)
    for i, line in enumerate(clean_lines):
        print(f"  [{i + 1}/{total}] {line}")
    print("=" * 60)

    # 2. 按完整历史事件分组，再合并过小的组（合并后不超过 12 条）
    raw_groups = group_outline_by_event(clean_lines)
    groups = merge_small_groups(raw_groups, max_size=12)

    print(f"\n按完整事件分组（并合并 ≤3 条的相邻小事件，合并上限 12 条），共 {len(groups)} 个文档：")
    print("-" * 60)
    for gi, g in enumerate(groups):
        print(f"  文档{gi + 1}（{len(g)} 条）: 行号 {g}")
    print("-" * 60)

    # 3. 串行扩写：每个事件组扩写完立即落盘一个 Word（中途崩了也保住已完成文档）
    print(f"\n第二步：串行扩写，每个事件组存一个文档，共 {len(groups)} 个文件...")
    prev_tail = ""  # 跨文档保持衔接，不因分文件而断
    for gi, group in enumerate(groups):
        print(f"\n===== 文档{gi + 1}/{len(groups)}（{len(group)} 条）=====")
        buffer = []
        for n in group:
            line = clean_lines[n - 1]
            print(f"\n  [行{n}/{total}] 正在扩写...")
            detail = generate_section_details(line, prev_tail)
            buffer.append(detail)
            prev_tail = get_tail(detail)
            print(f"    ✔ [行{n}] {line[:40]}... {len(detail)}字")
        # 本组扩写完，立即落盘
        chunk_md = "\n\n---\n\n".join(buffer)
        markdown_to_word(chunk_md, f"{TOPIC}_{gi + 1}")

    print("\n全流程结束！请去当前文件夹查看你的 Word 文件。")
