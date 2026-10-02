import statistics
import fitz
import pytesseract
from PIL import Image, ImageOps, ImageEnhance
from app.personnummer import find_candidates, score_candidate

class Analyzer:
    def __init__(self, language='swe+eng', dpi=300, image_ocr=True):
        self.language,self.dpi,self.image_ocr=language,dpi,image_ocr
    def analyze_page(self, page):
        embedded=page.get_text('text') or ''
        observations=[(raw,val,'embedded_text',95.0) for raw,val in find_candidates(embedded)]
        if not observations and self.image_ocr:
            observations.extend(self._image_ocr(page))
        grouped={}
        for raw,val,source,conf in observations: grouped.setdefault(val,[]).append((raw,source,conf))
        rows=[]
        for val,items in grouped.items():
            avg=statistics.mean(x[2] for x in items);score,reasons=score_candidate(val,len(items),avg)
            rows.append((items[0][0],val,','.join(sorted({x[1] for x in items})),avg,score,'; '.join(reasons)))
        rows.sort(key=lambda x:x[4],reverse=True)
        return embedded,rows
    def _image_ocr(self,page):
        zoom=self.dpi/72; pix=page.get_pixmap(matrix=fitz.Matrix(zoom,zoom),colorspace=fitz.csRGB,alpha=False)
        image=Image.frombytes('RGB',(pix.width,pix.height),pix.samples);gray=ImageOps.grayscale(image)
        variants={'gray':gray,'contrast':ImageEnhance.Contrast(gray).enhance(2.0),'threshold':gray.point(lambda p:255 if p>170 else 0)}
        found=[]
        for name,img in variants.items():
            data=pytesseract.image_to_data(img,lang=self.language,config='--psm 11',output_type=pytesseract.Output.DICT)
            text=' '.join(t for t in data['text'] if t.strip());confs=[float(c) for c in data['conf'] if str(c) not in ('-1','')];conf=sum(confs)/len(confs) if confs else 0
            found.extend((raw,val,'image_'+name,conf) for raw,val in find_candidates(text))
        return found
