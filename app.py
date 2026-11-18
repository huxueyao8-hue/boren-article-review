import io,os,csv,zipfile,re
import streamlit as st
from reviewer import review_document
st.set_page_config(page_title="博仁医院｜文章审核工作台",page_icon="📝",layout="wide")
st.title("📝 北京博仁医院｜文章审核工作台")
st.caption("部门共享网页版 · 批量 Word 审核 · 原文高亮 + Word 原生批注 · 无需登录")
st.info("本工具用于编辑初审，不代替医学、法务和品牌审核。默认规则模式不能核实复杂医学事实。")
def setting(key,default=""):
    try:return str(st.secrets.get(key,os.environ.get(key,default)))
    except Exception:return os.environ.get(key,default)
key=setting("AI_API_KEY")
with st.sidebar:
    st.header("审核设置")
    mode=st.radio("选择审核模式",["规则初审"]+(["AI 深度审核 + 规则"] if key else []))
    st.warning("公网服务请仅上传脱敏测试稿；含患者个人信息及未公开内部资料的文件必须先取得医院信息安全批准。")
files=st.file_uploader("批量上传 Word 文章（每批最多30篇）",type=["docx"],accept_multiple_files=True)
st.metric("本批文件",len(files) if files else 0)
if st.button("开始批量审核",type="primary",disabled=not files,use_container_width=True):
    if len(files)>30:st.error("每批最多30篇")
    elif any(f.size>10*1024*1024 for f in files):st.error("单文件不能超过10MB")
    else:
        output=io.BytesIO();summary=[]
        with zipfile.ZipFile(output,"w",zipfile.ZIP_DEFLATED) as archive:
            for i,f in enumerate(files):
                fname=os.path.basename(f.name).replace("/","_").replace("\","_")
                stem=re.sub(r"[^\w\u4e00-\u9fff. -]+","_",os.path.splitext(fname)[0]) or "文章"
                prefix=f"{i+1:02d}_{stem}"
                try:
                    marked,report,issues,count=review_document(f.getvalue(),fname,mode.startswith("AI"),key,setting("AI_BASE_URL","https://api.openai.com/v1"),setting("AI_MODEL","gpt-4.1-mini"))
                    archive.writestr(prefix+"_审核批注.docx",marked)
                    archive.writestr(prefix+"_审核报告.md",report)
                    summary.append({"文章":fname,"风险线索":len(issues),"已插入批注":count,"状态":"完成"})
                except Exception as exc:
                    summary.append({"文章":fname,"风险线索":"—","已插入批注":"—","状态":"失败："+type(exc).__name__})
            buff=io.StringIO();writer=csv.DictWriter(buff,fieldnames=["文章","风险线索","已插入批注","状态"]);writer.writeheader();writer.writerows(summary)
            archive.writestr("审核汇总.csv","\ufeff"+buff.getvalue())
        st.session_state["archive"]=output.getvalue()
        st.session_state["summary"]=summary
if "archive" in st.session_state:
    st.dataframe(st.session_state["summary"],hide_index=True,use_container_width=True)
    st.download_button("下载带批注 Word 和审核报告（ZIP）",st.session_state["archive"],file_name="博仁医院_审核结果.zip",mime="application/zip",type="primary",use_container_width=True)
st.caption("输出文件为新 Word 副本，不覆盖原文。部分超链接、文本框、页眉页脚和修订内容可能无法精确标注。")
