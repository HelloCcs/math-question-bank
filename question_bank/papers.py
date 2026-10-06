from __future__ import annotations
from dataclasses import dataclass

TYPE_ORDER={"单选题":0,"多选题":1,"填空题":2,"解答题":3}

@dataclass(frozen=True)
class PaperExportModel:
    title:str; exam_time:str; total_score:str; school:str=""; grade:str=""; class_name:str=""; instructions:str=""; questions:tuple=()

def build_export_model(paper:dict) -> PaperExportModel:
    missing=[q for q in paper.get("questions",[]) if q.get("missing")]
    if missing: raise ValueError("组卷中存在已删除题目，请先移除后再导出。")
    indexed=list(enumerate(paper.get("questions",[])))
    indexed.sort(key=lambda x:(TYPE_ORDER.get(x[1].get("question_type"),99),x[0]))
    questions=[]
    for number,(_,q) in enumerate(indexed,1):
        item=dict(q); item["export_number"]=number; questions.append(item)
    return PaperExportModel(*(str(paper.get(k) or "") for k in ("title","exam_time","total_score","school","grade","class_name","instructions")),tuple(questions))
