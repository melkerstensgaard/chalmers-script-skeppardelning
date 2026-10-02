import sys,json
from pathlib import Path
from PySide6.QtCore import Qt
from PySide6.QtWidgets import *
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from app.database import Database
from app.analyzer import Analyzer
from app.batch import BatchProcessor
from app.exporter import export_project
from PySide6.QtCore import QPointF


CLASSES={'Nytt bevis':'new_certificate','Tillhör föregående':'continuation','Provpapper':'exam_paper','Bilaga':'attachment','Annan handling':'other_document','Personnummer oläsligt':'unreadable','Osäker':'uncertain'}
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__();self.resize(1450,900);self.setWindowTitle('Sjöfartsbevis 0.3 - batchprojekt');self.db=self.project_id=self.volume_id=None;self.rows=[];self.pos=0
        self.doc=QPdfDocument(self);self.view=QPdfView();self.view.setDocument(self.doc);self.view.setPageMode(QPdfView.PageMode.SinglePage);self.view.setZoomMode(QPdfView.ZoomMode.FitInView)
        self.volume=QComboBox();self.volume.currentIndexChanged.connect(self.change_volume);self.info=QLabel('Inget projekt öppnat');self.pn=QLineEdit();self.score=QLabel('0/100');self.reasons=QTextEdit();self.reasons.setReadOnly(True);self.candidates=QComboBox();self.candidates.currentIndexChanged.connect(self.pick_candidate);self.cls=QComboBox();self.cls.addItems(CLASSES);self.comment=QTextEdit();self.comment.setMaximumHeight(80)
        panel=QWidget();form=QFormLayout(panel)
        for label,widget in [('Volym/PDF',self.volume),('Sida',self.info),('Kandidater',self.candidates),('Personnummer',self.pn),('Säkerhet',self.score),('Poängorsaker',self.reasons),('Klass',self.cls),('Kommentar',self.comment)]:form.addRow(label,widget)
        save=QPushButton('Spara och nästa ogranskade');save.clicked.connect(self.save_next);form.addRow(save);prev=QPushButton('Föregående');prev.clicked.connect(lambda:self.navigate(-1));nxt=QPushButton('Nästa');nxt.clicked.connect(lambda:self.navigate(1));nav=QHBoxLayout();nav.addWidget(prev);nav.addWidget(nxt);form.addRow(nav)
        center=QWidget();layout=QHBoxLayout(center);layout.addWidget(self.view,3);layout.addWidget(panel,1);self.setCentralWidget(center)
        menu=self.menuBar().addMenu('Projekt')
        for text,slot in [('Importera signummapp eller alla volymer',self.import_project),('Öppna arbetsdatabas',self.open_database),('Exportera alla färdiggranskade',self.export_all)]:menu.addAction(text).triggered.connect(slot)
    def import_project(self):
        root=QFileDialog.getExistingDirectory(self,'Välj en signummapp eller rotmappen med alla signummappar')
        if not root:return
        output=QFileDialog.getExistingDirectory(self,'Välj gemensam outputmapp')
        if not output:return
        workspace=QFileDialog.getExistingDirectory(self,'Välj arbetsmapp för databasen')
        if not workspace:return
        self.db=Database(Path(workspace)/'review.sqlite');processor=BatchProcessor(self.db,Analyzer());jobs=processor.discover(root)
        if not jobs:QMessageBox.warning(self,'Inga PDF-filer','Inga PDF-filer hittades.');return
        dialog=QProgressDialog('Analyserar materialet','Avbryt',0,len(jobs),self);dialog.setWindowModality(Qt.WindowModal)
        def progress(i,total,signum,name):dialog.setLabelText(f'{signum} / {name}');dialog.setValue(i);QApplication.processEvents()
        self.project_id,result=processor.run(root,output,progress,dialog.wasCanceled);dialog.close();self.populate_volumes();QMessageBox.information(self,'Importresultat',json.dumps(result,ensure_ascii=False,indent=2))
    def open_database(self):
        path,_=QFileDialog.getOpenFileName(self,'Öppna review.sqlite','','SQLite (*.sqlite *.db)')
        if not path:return
        self.db=Database(path);projects=self.db.projects()
        if not projects:QMessageBox.warning(self,'Tom databas','Inga projekt hittades.');return
        labels=[f"{p['input_root']} -> {p['output_root']}" for p in projects];choice,ok=QInputDialog.getItem(self,'Välj projekt','Projekt',labels,0,False)
        if ok:self.project_id=projects[labels.index(choice)]['id'];self.populate_volumes()
    def populate_volumes(self):
        self.volume.blockSignals(True);self.volume.clear()
        for v in self.db.volumes(self.project_id):self.volume.addItem(f"{v['signum']} | {v['source_name']} | {v['analysis_status']} | {v['analyzed_pages']}/{v['page_count']}",v['id'])
        self.volume.blockSignals(False)
        if self.volume.count():self.volume_id=self.volume.itemData(0);self.load_volume()
    def change_volume(self):
        if self.volume.currentData():self.volume_id=self.volume.currentData();self.load_volume()
    def load_volume(self):
        self.rows=list(self.db.pages(self.volume_id));self.pos=next((i for i,r in enumerate(self.rows) if r['review_status']!='reviewed'),0);v=next(x for x in self.db.volumes() if x['id']==self.volume_id);self.doc.load(v['source_path']);self.show_page()

    def show_page(self):
        if not self.rows:self.info.setText('Inga analyserade sidor');return
        r=self.rows[self.pos];self.view.pageNavigator().jump(
    int(r['page_index']),
    QPointF(0.0, 0.0),
    1.0
);self.info.setText(f"{r['page_index']+1}/{len(self.rows)} | {r['review_status']}");self.pn.setText(r['final_personnummer'] or r['suggested_personnummer'] or '');self.score.setText(f"{r['confidence_score']}/100");self.reasons.setPlainText(r['confidence_reasons']);self.comment.setPlainText(r['review_comment'] or '')
        self.candidates.clear();self.candidates.addItem('Välj kandidat','')
        for c in self.db.candidates(r['id']):self.candidates.addItem(f"{c['normalized_value']} ({c['score']})",c['normalized_value'])
        self.cls.setCurrentText(next((k for k,v in CLASSES.items() if v==r['page_class']),'Osäker'))

    def pick_candidate(self):
        if self.candidates.currentData():self.pn.setText(self.candidates.currentData())
    def save_next(self):
        if not self.rows:return
        r=self.rows[self.pos];self.db.review(r['id'],CLASSES[self.cls.currentText()],self.pn.text().strip(),self.comment.toPlainText().strip());self.rows=list(self.db.pages(self.volume_id));later=next((i for i,x in enumerate(self.rows) if i>self.pos and x['review_status']!='reviewed'),None)
        if later is not None:self.pos=later;self.show_page();return
        current=self.volume.currentIndex()
        if current+1<self.volume.count():self.volume.setCurrentIndex(current+1)
        else:self.show_page()
    def navigate(self,delta):self.pos=max(0,min(len(self.rows)-1,self.pos+delta));self.show_page()
    def export_all(self):
        done,failed=export_project(self.db,self.project_id)
        message=f"Exporterade PDF-volymer: {len(done)}\n\nEj exporterade:\n" + ('\n'.join(failed) if failed else 'Inga')
        QMessageBox.information(self,'Exportresultat',message)
def run_app():
    app=QApplication(sys.argv);window=MainWindow();window.show();sys.exit(app.exec())
