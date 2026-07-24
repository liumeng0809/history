import time
from openai import OpenAI
import httpx
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

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


def generate_section_details(outline_line):
    """第二步：针对大纲的每一行，生成详细正文"""
    prompt = f"""
    请针对这一历史节点：【{outline_line}】，写一篇非常详细的科普文章。
    要求：
    1. 字数不少于 800 字。
    2. 必须包含具体的年份、人物名字、事件起因、经过、结果和历史影响。
    3. 拒绝宏观叙事，不要泛泛而谈，只要具体史实细节。
    4. 直接输出正文内容，使用 Markdown 格式排版（用 ### 作为标题）。
    """
    print(f"  -> 正在详细扩写: {outline_line}")
    return call_ai(prompt)


def markdown_to_word(md_content, topic):
    """第三步：合并并转为 Word 文档"""
    print("\n第三阶段：正在生成 Word 文档...")

    doc = Document()

    # 设置正文中文字体（默认字体不含中文会显示异常）
    style = doc.styles['Normal']
    style.font.name = 'Microsoft YaHei'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
    style.font.size = Pt(12)

    # 大标题
    title = doc.add_heading(f"{topic}详细科普", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

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
                doc.add_paragraph(line)
        if sec.strip():
            doc.add_paragraph()  # 章节间空行

    word_file = f"{topic}_详细科普.docx"
    doc.save(word_file)
    print(f"\n🎉 恭喜！成功生成Word: {word_file}")


if __name__ == "__main__":
    # 1. 获取大纲
    outline = generate_outline(TOPIC)

    # 将大纲按行分割成列表，并过滤掉空行
    outline_lines = [line.strip() for line in outline.split('\n') if line.strip()]

    # 2. 遍历大纲，逐个扩写
    all_details = ""
    total = len(outline_lines)
    for i, line in enumerate(outline_lines):
        # 去掉大纲前面可能有的序号符号，让标题更干净
        clean_line = line.lstrip('0123456789.-* ')
        if not clean_line:
            continue

        print(f"\n[{i + 1}/{total}] 正在处理大纲第 {i + 1} 部分...")
        detail = generate_section_details(clean_line)
        all_details += f"{detail}\n\n---\n\n"  # 用分隔线分开每一节
        time.sleep(1)  # 稍微停顿1秒，防止调用API太快被限制

    # 3. 合并生成 Word
    markdown_to_word(all_details, TOPIC)
    print("\n全流程结束！请去当前文件夹查看你的 Word 文件。")
