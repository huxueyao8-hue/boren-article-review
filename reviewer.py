import io,re,json,urllib.request
from dataclasses import dataclass
from docx import Document
from docx.enum.text import WD_COLOR_INDEX

@dataclass
class Issue:
    quote:str
    category:str
    severity:str
    reason:str
    suggestion:str
    source:str="规则初审"

RULES=[
("绝对化疗效","高",r"彻底治愈|保证治愈|100%治愈|永不复发|绝不复发|零复发|无任何副作用|完全无风险|包治|一定能治好","不宜对疗效或风险作绝对承诺","删除绝对化措辞，交医学负责人复核"),
("首例排名","高",r"全球首例|世界首例|国内首例|全国首例|全球首台|全国首台|全国第一|世界第一|排名第[一二三四五六七八九十\\d]+|国际领先|国内领先|最先进|唯一一家","首例、排名和领先说法需要公开证据及限定范围","补充可核验来源、日期和统计口径，无法证明时改客观描述"),
("临床数据","中",r"\\d+(?:\\.\\d+)?\\s*[%％]|生存率|缓解率|治愈率|有效率|无病生存率|移植相关死亡率","临床数据需核验研究人群、样本、时点和证据","注明来源、样本量、观察期和适用人群"),
("就医承诺","高",r"无需排队|保证入院|保证床位|保证有效|医保全部报销|100%报销","存在无条件的服务或医保承诺","核查政策与适用范围，改为有条件的事实表达"),
("个人信息","高",r"身份证(?:号码|号)?[:：]?\\s*\\d{17}[\\dXx]|(?:手机|电话|联系方式)[:：]?\\s*1[3-9]\\d{9}|住址[:：]\\s*[^，。\\n]{5,30}","可能涉及可识别患者身份信息","核实患者授权及发布必要性，必要时脱敏"),
("病例授权","中",r"患者故事|真实案例|病友经历|患者自述|患儿家属|治疗前后对比|康复奇迹","患者案例需要核实授权及避免疗效暗示","核查授权、隐私处理和案例描述"),
("治疗术语","中",r"CAR[-－ ]?T|SCTIL|细胞治疗|临床试验|临床研究|靶向药|免疫治疗|新型疗法|脑深部电刺激|DBS","涉及专业治疗或新技术，应核验适应证与研究状态","请专家核对医学含义、适应证与证据"),
("合作背书","中",r"天坛医院|宣武医院|新华网|腾讯|国家级|权威认证|指定医院","涉及第三方合作、认证或背书","核查协议及宣传授权范围"),
("营销措辞","中",r"预约名额有限|抢占名额|立刻到院|马上咨询|最后机会|错过再等一年|一站式治愈|最便宜|最低价","可能存在诱导或不当比较","改为客观就诊信息"),
("医院名称","低",r"博仁医院|北京博仁医院|北京高博博仁医院","名称需按发布渠道与最新官方规范核验","检查医院全称、简称和首处出现方式")
]
def paras(doc):
    for p in doc.paragraphs:yield p
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                for p in c.paragraphs:yield p
def extract(doc):return "\n".join(p.text for p in paras(doc) if p.text.strip())
def review_rules(text):
    return [Issue(m.group(),c,s,r,a) for c,s,pat,r,a in RULES for m in list(re.finditer(pat,text,re.I))[:30]]
def review_ai(text,key,url,model):
    if not key:return []
    prompt="你是医院文章审核辅助。检查医学、广告合规、品牌、数据出处、语句逻辑、患者隐私。只能提出需复核的线索，不臆造事实。返回JSON对象：{\"issues\":[{\"quote\":\"文章中逐字存在的原句\",\"category\":\"分类\",\"severity\":\"高/中/低\",\"reason\":\"原因\",\"suggestion\":\"建议\"}]}。最多25条。"
    endpoint=url.rstrip("/")
    if not endpoint.endswith("/chat/completions"):endpoint+="/chat/completions"
    payload=json.dumps({"model":model,"temperature":0.1,"response_format":{"type":"json_object"},"messages":[{"role":"system","content":prompt},{"role":"user","content":text[:26000]}]},ensure_ascii=False).encode()
    req=urllib.request.Request(endpoint,data=payload,headers={"Authorization":"Bearer "+key,"Content-Type":"application/json"},method="POST")
    with urllib.request.urlopen(req,timeout=90) as resp: data=json.loads(resp.read().decode())
    output=json.loads(data["choices"][0]["message"]["content"])
    issues=[]
    for x in output.get("issues",[])[:25]:
        q=str(x.get("quote","")).strip()
        if not q or q not in text:continue
        sev=x.get("severity","中")
        issues.append(Issue(q,str(x.get("category","其他"))[:35],sev if sev in ("高","中","低") else "中",str(x.get("reason","需核验"))[:500],str(x.get("suggestion","交人工审核"))[:600],"AI辅助"))
    return issues
def split(p,offset):
    from copy import deepcopy
    total=0
    for run in list(p.runs):
        n=len(run.text)
        if total<offset<total+n:
            old=run.text;new=deepcopy(run._r);run._r.addnext(new)
            run.text=old[:offset-total]
            type(run)(new,p).text=old[offset-total:]
            return
        total+=n
def annotate(doc,issues):
    done=set();count=0
    for issue in sorted(issues,key=lambda x:len(x.quote),reverse=True):
        for pi,p in enumerate(paras(doc)):
            text=p.text;start=text.find(issue.quote)
            if start<0:continue
            end=start+len(issue.quote)
            if (pi,start,end) in done or "".join(r.text for r in p.runs)!=text:continue
            split(p,end);split(p,start);hits=[];pos=0
            for run in p.runs:
                nex=pos+len(run.text)
                if pos>=start and nex<=end and run.text:hits.append(run)
                pos=nex
            if not hits:continue
            color=WD_COLOR_INDEX.YELLOW if issue.severity=="高" else WD_COLOR_INDEX.TURQUOISE if issue.severity=="中" else WD_COLOR_INDEX.GRAY_25
            for run in hits:run.font.highlight_color=color
            doc.add_comment(hits,text=f"【{issue.severity}风险｜{issue.category}｜{issue.source}】\n原因：{issue.reason}\n建议：{issue.suggestion}\n请由相关负责人确认，不构成最终审核结论。",author="博仁内容审核助手",initials="AI")
            done.add((pi,start,end));count+=1;break
    return count
def review_document(blob,name,enable_ai=False,key="",base_url="https://api.openai.com/v1",model="gpt-4.1-mini"):
    doc=Document(io.BytesIO(blob));text=extract(doc)
    if not text.strip():raise ValueError("未提取到可编辑文本")
    issues=review_rules(text)
    if enable_ai:issues+=review_ai(text,key,base_url,model)
    distinct=[];seen=set()
    for x in issues:
        k=(x.quote,x.category,x.reason)
        if k not in seen:distinct.append(x);seen.add(k)
    marked=annotate(doc,distinct)
    out=io.BytesIO();doc.save(out)
    report=["# 审核报告："+name,"","共识别 "+str(len(distinct))+" 条风险线索；已插入 "+str(marked)+" 条Word批注。","仅用于内容初筛，医学数据、合规判断和患者授权均需人工复核。",""]
    for i,x in enumerate(distinct,1):
        report += [f"## {i}. [{x.severity}] {x.category}（{x.source}）",f"- 原文：{x.quote}",f"- 原因：{x.reason}",f"- 建议：{x.suggestion}",""]
    return out.getvalue(),"\n".join(report),distinct,marked
