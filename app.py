import streamlit as st
import os
import time
import concurrent.futures
from generate_review import (
    extract_topics,
    generate_topic_details,
    generate_mindmap,
    generate_connection_text,
    generate_formula_summary,
    extract_text_from_pptx,
    extract_text_from_pdf
)

st.set_page_config(page_title="复习资料生成器", layout="wide")
st.title("📚 复习资料生成器")

last_result_file = "last_result.md"

# ========== 显示上次结果（如果存在） ==========
if os.path.exists(last_result_file):
    with open(last_result_file, "r", encoding="utf-8") as f:
        last_content = f.read()
    with st.sidebar.expander("上次生成结果", expanded=False):
        st.download_button(
            label="下载上次结果",
            data=last_content.encode("utf-8"),
            file_name="last_result.md",
            mime="text/markdown"
        )
        if st.button("显示上次结果"):
            st.markdown(last_content)

# ========== 输入方式选择 ==========
input_method = st.radio(
    "选择内容输入方式",
    ["粘贴文本", "上传 TXT", "上传 PPTX", "上传 PDF"]
)

chapter_text = ""

if input_method == "粘贴文本":
    chapter_text = st.text_area("请粘贴章节内容", height=300)
elif input_method == "上传 TXT":
    uploaded_file = st.file_uploader("上传 TXT 文件", type=["txt"])
    if uploaded_file is not None:
        chapter_text = uploaded_file.read().decode("utf-8")
elif input_method == "上传 PPTX":
    uploaded_file = st.file_uploader("上传 PPTX 文件", type=["pptx"])
    if uploaded_file is not None:
        with open("temp.pptx", "wb") as f:
            f.write(uploaded_file.getbuffer())
        chapter_text = extract_text_from_pptx("temp.pptx")
elif input_method == "上传 PDF":
    uploaded_file = st.file_uploader("上传 PDF 文件", type=["pdf"])
    if uploaded_file is not None:
        with open("temp.pdf", "wb") as f:
            f.write(uploaded_file.getbuffer())
        chapter_text = extract_text_from_pdf("temp.pdf")

# ========== 题型数量设置 ==========
st.subheader("题型数量设置")
col1, col2, col3, col4 = st.columns(4)
with col1:
    num_choice = st.number_input("选择题", 0, 10, 2, key="choice")
with col2:
    num_fill = st.number_input("填空题", 0, 10, 0, key="fill")
with col3:
    num_short = st.number_input("简答题", 0, 10, 2, key="short")
with col4:
    num_calc = st.number_input("计算题", 0, 10, 0, key="calc")

exam_config = {
    "multiple_choice": num_choice,
    "fill_blank": num_fill,
    "short_answer": num_short,
    "calculation": num_calc
}

# ========== 缓存生成函数 ==========
@st.cache_data(show_spinner=False)
def cached_generate_review(chapter_text, exam_config_tuple):
    """
    缓存生成结果，输入完全相同时直接返回缓存内容。
    包含进度条和状态提示。
    """
    exam_config = dict(exam_config_tuple)
    topics = extract_topics(chapter_text)
    if not topics:
        return None

    total = len(topics)
    progress_bar = st.progress(0)
    status_text = st.empty()

    # 并行生成每个考点的复习卡片
    def generate_one(topic):
        return topic, generate_topic_details(
            topic,
            chapter_text[:600],
            exam_config,
            topics
        )

    results = []
    max_workers = 3   # 可调整并发数

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_topic = {executor.submit(generate_one, topic): topic for topic in topics}
        completed = 0
        for future in concurrent.futures.as_completed(future_to_topic):
            try:
                topic, detail = future.result()
                results.append((topic, detail))
            except Exception as e:
                st.error(f"生成失败：{e}")
            completed += 1
            progress_bar.progress(completed / total)
            status_text.text(f"正在生成考点卡片... 已完成 {completed}/{total}")

    # 生成知识框架（列表形式）
    status_text.text("正在生成知识框架...")
    mindmap_text = generate_mindmap(topics)

    # 生成知识点串联
    status_text.text("正在生成知识点串联...")
    connection_text = generate_connection_text(topics)

    # 生成公式总结
    status_text.text("正在生成公式总结...")
    formula_summary = generate_formula_summary(chapter_text, topics)

    # 清理进度条和状态
    progress_bar.empty()
    status_text.empty()

    results.sort(key=lambda x: x[0]['id'])
    all_details = [detail for _, detail in results]

    # 组合：知识框架（Markdown列表）+ 串联 + 公式总结
    combined_framework = (
        "## 知识框架\n" + mindmap_text + "\n\n"
        "## 知识点串联\n" + connection_text + "\n\n"
        "## 公式总结\n" + formula_summary
    )

    return combined_framework, all_details

# ========== 生成按钮 ==========
if st.button("生成复习资料"):
    if not chapter_text.strip():
        st.error("请先输入或上传章节内容")
    else:
        with st.spinner("正在生成，请稍候..."):
            # 预估等待时间（粗略）
            estimated_time = 20 * 5  # 假设平均5个考点，每个20秒
            st.info(f"预计等待时间约 {estimated_time} 秒，实际取决于模型响应速度")
            
            exam_config_tuple = tuple(sorted(exam_config.items()))
            result = cached_generate_review(chapter_text, exam_config_tuple)
            if result is None:
                st.error("考点提取失败")
            else:
                combined_framework, all_details = result
                st.success(f"成功生成 {len(all_details)} 个考点")

                # 显示知识框架+串联+公式总结
                st.markdown(combined_framework)

                # 显示核心考点复习卡片
                st.subheader("核心考点复习卡片")
                for detail in all_details:
                    st.markdown(detail)
                    st.markdown("---")

                # 构造完整 Markdown 内容
                md_content = "# 复习资料\n\n" + combined_framework + "\n\n## 核心考点复习卡片\n"
                for detail in all_details:
                    md_content += detail + "\n---\n"

                # 保存到文件
                with open(last_result_file, "w", encoding="utf-8") as f:
                    f.write(md_content)

                st.download_button(
                    label="下载 Markdown 文件",
                    data=md_content.encode("utf-8"),
                    file_name="review_output.md",
                    mime="text/markdown"
                )