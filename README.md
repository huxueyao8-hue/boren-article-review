# 北京博仁医院文章智能审核工作台

免登录的 Streamlit 批量 DOCX 初审网页。支持规则审核、可选 AI 深度审核，以及导出含 Word 原生批注的文件。

**安全提醒：** 开源代码不包含任何患者数据和 API Key。公开部署的网站可被知道网址的人访问，不建议上传真实患者资料或未公开医院信息。正式使用前请完成医院信息安全审批。

部署：Render Python Web Service，Build Command `pip install -r requirements.txt`，Start Command `streamlit run app.py --server.address 0.0.0.0 --server.port $PORT --server.headless true`。

基础规则不能替代医学审核和法务审核。AI 功能需要管理员在 Render 环境变量中设置 `AI_API_KEY`，可选 `AI_BASE_URL` 和 `AI_MODEL`。
