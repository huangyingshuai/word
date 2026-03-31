"""
模块名称：main.py
模块职责：Streamlit界面渲染
"""
import streamlit as st
import pandas as pd
import json
from datetime import datetime
import copy

from config.templates import TEMPLATE_LIBRARY, apply_template_to_config, validate_template, DEFAULT_TEMPLATE
from config.constants import FONT_LIST, FONT_SIZE_LIST, ALIGN_LIST, LINE_TYPE_LIST, LINE_RULE, EN_FONT_LIST, FONT_SIZE_NUM
from services.doc_process_service import process_document
from services.test_service import run_standard_test
from core.number_recognizer import NUMBER_TYPE_DEF
from core.number_grouper import group_number_items, check_gb_compliance

# ====================== 页面配置 ======================
st.set_page_config(page_title="Word智能排版工具", layout="wide", page_icon="📄")

# ====================== 会话初始化 ======================
def init_session():
    if "current_config" not in st.session_state:
        st.session_state.current_config = copy.deepcopy(DEFAULT_TEMPLATE)
    if "template_version" not in st.session_state:
        st.session_state.template_version = 0
    if "title_records" not in st.session_state:
        st.session_state.title_records = []
    if "number_groups" not in st.session_state:
        st.session_state.number_groups = []
    if "last_template" not in st.session_state:
        st.session_state.last_template = "默认通用格式"
    if "corrected_number_items" not in st.session_state:
        st.session_state.corrected_number_items = None
    if "correction_history" not in st.session_state:
        st.session_state.correction_history = []
    if "group_format_config" not in st.session_state:
        st.session_state.group_format_config = {}

init_session()

# ====================== 主逻辑 ======================
def main():
    st.title("📄 Word智能排版工具")
    st.success("✅ 四层标题识别 | 序号统一转换 | 禁用自动编号 | 内容全面保护")

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📁 文档排版", "🔧 序号修正", "📊 识别预览", "🧪 准确率测试", "⚙️ 模板管理"])

    # Tab1 文档排版
    with tab1:
        st.subheader("Step 1: 选择模板")
        keep_custom = st.checkbox("保留自定义格式", False)
        tpl1, tpl2, tpl3 = st.tabs(["高校论文", "通用办公", "党政公文"])
        with tpl1:
            t = st.selectbox("高校模板", [k for k in TEMPLATE_LIBRARY if "河北" in k or "国标" in k])
            if st.button("应用高校模板", use_container_width=True):
                st.session_state.current_config = apply_template_to_config(t, keep_custom, st.session_state.current_config)
                st.success(f"应用【{t}】成功")
        with tpl2:
            t = st.selectbox("通用模板", [k for k in TEMPLATE_LIBRARY if "默认" in k])
            if st.button("应用通用模板", use_container_width=True):
                st.session_state.current_config = apply_template_to_config(t, keep_custom, st.session_state.current_config)
                st.success(f"应用【{t}】成功")
        with tpl3:
            t = st.selectbox("公文模板", [k for k in TEMPLATE_LIBRARY if "党政" in k])
            if st.button("应用公文模板", use_container_width=True):
                st.session_state.current_config = apply_template_to_config(t, keep_custom, st.session_state.current_config)
                st.success(f"应用【{t}】成功")

        # 自定义格式
        st.subheader("Step 2: 自定义格式")
        cfg = st.session_state.current_config
        tabs = st.tabs(["一级标题", "二级标题", "三级标题", "正文"])
        for i, level in enumerate(["一级标题", "二级标题", "三级标题", "正文"]):
            with tabs[i]:
                c1,c2,c3,c4 = st.columns(4)
                with c1:
                    cfg[level]["font"] = st.selectbox("字体", FONT_LIST, FONT_LIST.index(cfg[level]["font"]), key=f"{level}_font")
                    cfg[level]["size"] = st.selectbox("字号", FONT_SIZE_LIST, FONT_SIZE_LIST.index(cfg[level]["size"]), key=f"{level}_size")
                with c2:
                    cfg[level]["bold"] = st.checkbox("加粗", cfg[level]["bold"], key=f"{level}_bold")
                    cfg[level]["align"] = st.selectbox("对齐", ALIGN_LIST, ALIGN_LIST.index(cfg[level]["align"]), key=f"{level}_align")
                with c3:
                    cfg[level]["line_type"] = st.selectbox("行距", LINE_TYPE_LIST, LINE_TYPE_LIST.index(cfg[level]["line_type"]), key=f"{level}_line")
                    rule = LINE_RULE[cfg[level]["line_type"]]
                    cfg[level]["line_value"] = st.slider(rule["label"], rule["min"], rule["max"], cfg[level]["line_value"], key=f"{level}_val")
                with c4:
                    if level != "正文":
                        cfg[level]["space_before"] = st.number_input("段前", 0.0, 100.0, cfg[level]["space_before"], key=f"{level}_b")
                        cfg[level]["space_after"] = st.number_input("段后", 0.0, 100.0, cfg[level]["space_after"], key=f"{level}_a")
                    cfg[level]["indent"] = st.number_input("缩进", 0,10,cfg[level]["indent"], key=f"{level}_in")

        # 数字格式
        st.subheader("正文数字/英文格式")
        num_cfg = {"enable": st.checkbox("启用单独格式", True)}
        nc1,nc2 = st.columns(2)
        with nc1:
            num_cfg["font"] = st.selectbox("字体", EN_FONT_LIST)
            num_cfg["size_same_as_body"] = st.checkbox("字号同正文", True)
        with nc2:
            num_cfg["size"] = st.selectbox("字号", FONT_SIZE_LIST) if not num_cfg["size_same_as_body"] else cfg["正文"]["size"]
            num_cfg["bold"] = st.checkbox("加粗", False)

        # 【新增】序号格式选择
        st.subheader("序号统一格式设置")
        target_num_format = st.selectbox(
            "选择序号统一格式",
            ["1.", "(1)", "①", "一、"],
            index=0,
            help="识别到的序号将统一转换为此格式"
        )

        # 高级选项
        st.subheader("Step 3: 高级选项")
        c1,c2,c3 = st.columns(3)
        with c1: 
            reg = st.checkbox("启用标题正则识别", True)
            disable_auto = st.checkbox("禁用Word自动编号", True, help="彻底防止重复生成序号")
        with c2: 
            force = st.checkbox("绑定Word内置样式", True)
            keep_spacing = st.checkbox("保留原有段落间距", False)
        with c3: 
            clear = st.checkbox("清理多余空行", True)
            max_blank = st.number_input("最大连续空行数", 1, 10, 1)

        # 文件处理
        st.subheader("Step 4: 上传并排版")
        file = st.file_uploader("上传docx", type="docx")
        if file and st.button("🚀 开始排版", type="primary", use_container_width=True):
            with st.spinner("正在处理文档，请稍候..."):
                res, stats, groups = process_document(
                    file.getvalue(), cfg, num_cfg,
                    enable_title_regex=reg,
                    enable_context_check=True,
                    force_style=force,
                    keep_spacing=keep_spacing,
                    clear_blank=clear,
                    max_blank=max_blank,
                    disable_auto_numbering=disable_auto,
                    target_number_format=target_num_format
                )
                if res:
                    st.success("✅ 文档排版完成！")
                    st.session_state.number_groups = groups
                    st.download_button("📥 下载排版后的文档", res, f"排版后_{file.name}", use_container_width=True)
                    c1,c2,c3,c4,c5 = st.columns(5)
                    c1.metric("一级标题", stats["一级标题"])
                    c2.metric("二级标题", stats["二级标题"])
                    c3.metric("三级标题", stats["三级标题"])
                    c4.metric("序号项", stats["序号项"])
                    c5.metric("图片/表格", f"{stats['图片']}/{stats['表格']}")

    # Tab2 序号修正
    with tab2:
        st.subheader("序号分组与格式")
        if not st.session_state.number_groups:
            st.warning("请先在「文档排版」页上传并处理文档")
        else:
            for g in st.session_state.number_groups:
                with st.expander(f"序号组 {g['group_id']} | {NUMBER_TYPE_DEF[g['type']]['gb_name']} | {g['level']}级 | 共{len(g['items'])}项"):
                    ok, issues = check_gb_compliance(g["type"], g["level"])
                    if ok:
                        st.success("✅ 符合国标GB/T 1.1-2020规范")
                    else:
                        for issue in issues:
                            st.warning(f"⚠️ {issue}")
                    
                    st.subheader("序号项明细")
                    for item in g["items"]:
                        st.write(f"段落 {int(item['para_index'])+1}：{item['full_text']}")

    # Tab3 预览
    with tab3:
        st.subheader("标题识别结果预览")
        if st.session_state.title_records:
            df = pd.DataFrame(st.session_state.title_records)
            st.dataframe(df, use_container_width=True)
        else:
            st.info("上传并处理文档后，可在此查看标题识别明细")

    # Tab4 测试
    with tab4:
        st.subheader("序号识别准确率标准测试")
        if st.button("运行标准测试用例", use_container_width=True):
            with st.spinner("正在测试..."):
                acc, err = run_standard_test()
                st.metric("标准测试集准确率", f"{acc}%")
                if err:
                    st.subheader("错误用例")
                    for case in err:
                        st.write(f"测试文本：{case['text']}")
                        st.write(f"预期：类型={case['expected']['type']}，层级={case['expected']['level']}")
                        st.write(f"实际：类型={case['actual']['type']}，层级={case['actual']['level']}")
                        st.divider()
                else:
                    st.success("✅ 所有测试用例全部通过！")

    # Tab5 模板
    with tab5:
        st.subheader("模板管理")
        st.download_button("导出当前配置为模板", json.dumps(cfg, ensure_ascii=False, indent=2), "自定义模板.json")
        
        st.subheader("导入自定义模板")
        uploaded_tpl = st.file_uploader("上传模板配置文件(.json)", type=["json"])
        if uploaded_tpl:
            try:
                import_tpl = json.load(uploaded_tpl)
                is_valid, msg = validate_template(import_tpl)
                if is_valid:
                    st.success("✅ 模板校验通过")
                    if st.button("应用导入的模板", use_container_width=True):
                        st.session_state.current_config = import_tpl
                        st.session_state.template_version += 1
                        st.success("✅ 已应用导入的模板")
                else:
                    st.error(f"❌ 模板校验失败：{msg}")
            except Exception as e:
                st.error(f"❌ 模板文件解析失败：{str(e)}")

if __name__ == "__main__":
    main()