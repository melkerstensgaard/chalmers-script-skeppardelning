from pathlib import Path
from pypdf import PdfReader,PdfWriter
import csv,re

def safe(value): return re.sub(r'[^0-9A-Za-zÅÄÖåäö_-]+','_',value or 'SAKNAS')
def write_pages(reader,indices,path):
    writer=PdfWriter()
    for i in indices: writer.add_page(reader.pages[i])
    with open(path,'wb') as f: writer.write(f)
def export_volume(db,volume,target):
    pages=db.pages(volume['id']);errors=[];current_exists=False
    for p in pages:
        if p['review_status']!='reviewed' or p['page_class']=='uncertain':errors.append(f"Sida {p['page_index']+1} är inte slutgranskad")
        if p['page_class'] in ('new_certificate','unreadable'):current_exists=True
        elif p['page_class']=='continuation' and not current_exists:errors.append(f"Sida {p['page_index']+1} är fortsättning utan föregående bevis")
    if errors: raise ValueError('\n'.join(errors))
    out=Path(target);cert=out/'utbildningsbevis';cert.mkdir(parents=True,exist_ok=True);reader=PdfReader(volume['source_path']);docs=[];cur=None
    folders={'exam_paper':'provpapper','attachment':'bilagor','other_document':'andra_handlingar'}
    for p in pages:
        cls=p['page_class']
        if cls in ('new_certificate','unreadable'):
            if cur:docs.append(cur)
            cur={'pages':[p['page_index']],'pn':p['final_personnummer'],'score':p['confidence_score'],'status':cls}
        elif cls=='continuation':cur['pages'].append(p['page_index'])
        elif cls in folders:
            d=out/folders[cls];d.mkdir(parents=True,exist_ok=True);write_pages(reader,[p['page_index']],d/f"{safe(volume['signum'])}_sida_{p['page_index']+1:06d}.pdf")
    if cur:docs.append(cur)
    rows=[]
    for n,d in enumerate(docs,1):
        marker=d['pn'] or ('OLASLIGT' if d['status']=='unreadable' else 'SAKNAS');name=f"{safe(volume['signum'])}_{n:06d}_{safe(marker)}.pdf";write_pages(reader,d['pages'],cert/name)
        rows.append([name,volume['signum'],volume['source_name'],d['pages'][0]+1,d['pages'][-1]+1,d['pn'] or '',d['score']])
    with open(out/'index.csv','w',newline='',encoding='utf-8-sig') as f:csv.writer(f).writerows([['filename','signum','source_filename','start_page','end_page','personnummer','confidence_score'],*rows])
    db.set_exported(volume['id']);return rows
def export_project(db,project_id):
    project=db.project(project_id);done=[];failed=[]
    for volume in db.volumes(project_id):
        target=Path(project['output_root'])/safe(volume['signum'])/safe(Path(volume['source_name']).stem)
        try:export_volume(db,volume,target);done.append(volume['source_name'])
        except Exception as exc:failed.append(f"{volume['signum']} / {volume['source_name']}: {exc}")
    return done,failed
