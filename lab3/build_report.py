import docx, json, os
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH as A
from docx.oxml.ns import qn

TPL="/root/.claude/uploads/5b43f2ec-e0ff-576e-8e16-7d5d355ff10e/f1d8673c-lab-1_report.docx"
HERE=os.path.dirname(os.path.abspath(__file__))
OUT=os.path.join(HERE,"lab-3_report.docx")
d=docx.Document(TPL)
body=d.element.body
for el in list(body)[12:]:
    if el.tag!=qn('w:sectPr'): body.remove(el)
P=d.paragraphs
P[2].runs[0].text="Лабораторная работа № 3"
P[3].runs[0].text="Мониторинг сетевого трафика на хосте на примере работы с утилитами диагностики и мониторинга сетевых соединений в Linux"
for r in P[3].runs[1:]: r.text=""

def para(text="",size=11,bold=False,font="Times New Roman",style=None,align=None):
    p=d.add_paragraph(style=style)
    if text:
        r=p.add_run(text); r.font.name=font; r.font.size=Pt(size); r.bold=bold
        r._r.rPr.rFonts.set(qn('w:eastAsia'),font)
    if align: p.alignment=align
    return p
def code(t):
    for line in t.strip("\n").split("\n"):
        p=para(line,8.5,font="Courier New"); p.paragraph_format.space_after=Pt(0)
def shot(path,caption):
    p=para(align=A.CENTER); p.add_run().add_picture(os.path.join(HERE,path),width=Cm(15))
    para(caption,10,align=A.CENTER)

# report.json: {"steps":[{"part":..,"text":..}], "artifacts":[{"h","text","code","img","cap"}], "answers":[{"q","a"}], "conclusion":".."}
R=json.load(open(os.path.join(HERE,"report.json")))
d.add_paragraph("Ход работы",style="Heading 1")
for s in R.get("steps",[]):
    para(s["part"],12,True)
    for t in s["text"]: para(t)
    para()
d.add_paragraph("Артефакты выполнения",style="Heading 1")
for a in R.get("artifacts",[]):
    if a.get("h"): d.add_paragraph(a["h"],style="Heading 2")
    if a.get("text"): para(a["text"])
    if a.get("code"): code(a["code"])
    for im in a.get("imgs",[]): shot(im["img"],im["cap"])
if R.get("answers"):
    d.add_paragraph("Ответы на вопросы и задания",style="Heading 1")
    for x in sorted(R["answers"],key=lambda x:x["n"]):
        para(f"{x['n']}. {x['q']}",11,True); para(x["a"])
if R.get("conclusion"):
    d.add_paragraph("Вывод",style="Heading 1"); para(R["conclusion"])
d.save(OUT); print("saved",OUT)
