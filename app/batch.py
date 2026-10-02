from pathlib import Path
import hashlib
import fitz
from pypdf import PdfReader

class BatchProcessor:
    def __init__(self,db,analyzer): self.db,self.analyzer=db,analyzer
    def discover(self,root):
        root=Path(root).resolve();jobs=[]
        direct=sorted(root.glob('*.pdf'))+sorted(root.glob('*.PDF'))
        jobs.extend((root.name,p) for p in direct)
        for folder in sorted(p for p in root.iterdir() if p.is_dir()):
            pdfs=sorted({*folder.rglob('*.pdf'),*folder.rglob('*.PDF')})
            jobs.extend((folder.name,p) for p in pdfs)
        return jobs
    @staticmethod
    def sha256(path):
        h=hashlib.sha256()
        with open(path,'rb') as f:
            for chunk in iter(lambda:f.read(1024*1024),b''): h.update(chunk)
        return h.hexdigest()
    def run(self,input_root,output_root,progress=None,cancelled=None):
        pid=self.db.add_project(input_root,output_root);jobs=self.discover(input_root)
        result={'found':len(jobs),'analyzed':0,'resumed':0,'skipped':0,'failed':0,'errors':[]}
        for job_no,(signum,path) in enumerate(jobs,1):
            if cancelled and cancelled(): break
            volume_id=None
            try:
                count=len(PdfReader(str(path)).pages);volume_id=self.db.add_volume(pid,signum,path,count,self.sha256(path))
                existing=len(self.db.pages(volume_id))
                if existing==count:
                    self.db.set_volume_status(volume_id,'analyzed',count);result['skipped']+=1
                else:
                    if existing: result['resumed']+=1
                    self.db.set_volume_status(volume_id,'analyzing',existing)
                    doc=fitz.open(path)
                    for page_index in range(existing,len(doc)):
                        if cancelled and cancelled(): break
                        text,candidates=self.analyzer.analyze_page(doc[page_index]);best=candidates[0] if candidates else None
                        page_id=self.db.save_page(volume_id,page_index,text,best[1] if best else None,best[4] if best else 0,best[5] if best else 'Ingen kandidat hittades')
                        self.db.replace_candidates(page_id,candidates);self.db.set_volume_status(volume_id,'analyzing',page_index+1)
                    doc.close();done=len(self.db.pages(volume_id))
                    self.db.set_volume_status(volume_id,'analyzed' if done==count else 'pending',done)
                    if done==count: result['analyzed']+=1
            except Exception as exc:
                result['failed']+=1;result['errors'].append(f'{path}: {exc}')
                if volume_id:self.db.set_volume_status(volume_id,'failed',error=str(exc))
            if progress: progress(job_no,len(jobs),signum,path.name)
        return pid,result
