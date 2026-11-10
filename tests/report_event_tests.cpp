#include "access_form_preview.hpp"
#include "original_bound_value.hpp"
#include "access_form_objects.hpp"
#include <QLineEdit>
#include <QApplication>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonObject>
#include <QCryptographicHash>
#include <QRegularExpression>
#include <QPushButton>
#include "vba_source_codec.hpp"
#include "access_query.hpp"
#include <QSqlRecord>
#include <QMainWindow>
#include <QMenuBar>
#include <QMenu>
#include <iostream>
#include <cmath>
#include <QToolBar>
#include <QComboBox>
#include <QSqlQuery>
#include <QTemporaryDir>
#include <QProcess>
#include <QStandardPaths>
#include <QPrintPreviewDialog>
#include <QPrintDialog>
#include <QPageLayout>
#include <QDialogButtonBox>
#include <QPrinter>
#include <QGraphicsView>
#include <QTest>
#include <QPrintPreviewWidget>
#include <QTimer>
#include <QAction>
#include <QTabWidget>
#include <QSpinBox>
#include <QDoubleSpinBox>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>

void testOpenReport();

void testReportEvents() {
    testOpenReport();
    auto db = QSqlDatabase::addDatabase("QSQLITE","report-event-tests");
    db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,name TEXT)"));
    assert(query.exec("INSERT INTO records VALUES(1,'first'),(2,'second')"));
    assert(query.exec("CREATE TABLE events(id INTEGER PRIMARY KEY,name TEXT)"));
    assert(query.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    query.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES('query','ranked',?)");
    query.addBindValue(QByteArray("SELECT TOP 1 id,name FROM records ORDER BY id")); assert(query.exec()); query.finish();
    const QByteArray source = R"REPORT(Begin Report
Caption ="Report title"
RecordSource ="ranked"
OnOpen ="[Event Procedure]"
OnClose ="[Event Procedure]"
OnNoData ="[Event Procedure]"
Begin Section
Height =1500
Begin TextBox
Name ="Name"
ControlSource ="name"
End
Begin TextBox
Name ="Info"
Top =400
End
End
End
)REPORT";
    const QByteArray code = R"VBA(
Private retainedReport As Report
Sub Report_Open(Cancel As Integer)
Dim typedReport As Access.Report
Set typedReport=Me
Set retainedReport=typedReport
If Not TypeOf Me Is Access.Report Then Err.Raise 5
If TypeOf Me Is Access.Form Then Err.Raise 5
If Not TypeOf Reports Is Access.Reports Then Err.Raise 5
CurrentDb.Execute "INSERT INTO events(name) VALUES('open')"
Me!Info = CStr(Reports.Count) & ":" & CStr(Forms.Count)
End Sub
Sub Report_NoData(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events(name) VALUES('nodata')"
Cancel = 0
End Sub
Sub Report_Close()
CurrentDb.Execute "INSERT INTO events(name) VALUES('close')"
End Sub
)VBA";
    const auto trace = [&] {
        QStringList rows;
        assert(query.exec("SELECT name FROM events ORDER BY id"));
        while (query.next()) rows.append(query.value(0).toString());
        query.finish(); return rows;
    };
    const auto clear = [&] { assert(query.exec("DELETE FROM events")); query.finish(); };
    auto bandedSource=source;
    bandedSource.replace("Begin Section\nHeight =1500",R"BANDS(Begin FormHeader
Height =300
BackColor =65280
End
Begin PageHeader
Height =600
BackColor =16711680
Begin TextBox
Name ="PageTitle"
Top =150
End
End
Begin PageFooter
Height =450
BackColor =255
Begin TextBox
Name ="PageNumber"
Top =150
End
End
Begin FormFooter
Height =300
Begin TextBox
Name ="ReportSummary"
Top =150
End
End
Begin Section
Height =1500)BANDS");
    const auto bandedLayout=linaccess::parseAccessFormLayout(QString::fromUtf8(bandedSource));
    assert(bandedLayout.sectionHeights.value("PageHeader")==600);
    assert(bandedLayout.sectionHeights.value("PageFooter")==450);
    assert(bandedLayout.sectionOffsets.value("Section")==900);
    assert(bandedLayout.sectionOffsets.value("FormFooter")==2400);
    assert(bandedLayout.sectionOffsets.value("PageFooter")==2700 && bandedLayout.height==3150);
    {
        AccessFormPreview banded("banded report",bandedSource,{},db);
        const auto field=[&](const char* name) {
            auto* widget=banded.findChild<QLineEdit*>(QString("access-control-")+name);
            assert(widget); return widget;
        };
        assert(field("Name")->text()=="first" && field("Name")->isReadOnly());
        assert(field("PageTitle")->y()==30 && field("Name")->y()==60);
        assert(field("ReportSummary")->y()==170 && field("PageNumber")->y()==190);
        auto* canvas=banded.findChild<QWidget*>("access-form-canvas"); assert(canvas);
        const auto colors=canvas->grab().toImage();
        assert(colors.pixelColor(200,10)==QColor(Qt::green));
        assert(colors.pixelColor(200,30)==QColor(Qt::blue));
        assert(colors.pixelColor(200,100)==QColor(Qt::white));
        assert(colors.pixelColor(200,200)==QColor(Qt::red));

        QTemporaryDir output; assert(output.isValid());
        const auto single=output.filePath("single.pdf");
        // The summary extends outside its short section: do not silently clip it.
        assert(!banded.exportReportPdf(single));
        assert(field("Name")->text()=="first");
        auto pagedSource=bandedSource;
        pagedSource.replace("RecordSource =\"ranked\"","RecordSource =\"records\"");
        pagedSource.replace("Height =1500","Height =9000");
        pagedSource.replace("Height =300","Height =600");
        pagedSource.replace("Name =\"Info\"\nTop", "Name =\"Info\"\nControlSource =\"=[id] * 10\"\nTop");
        pagedSource.replace("Name =\"ReportSummary\"\nTop", "Name =\"ReportSummary\"\nControlSource =\"=Sum([id])\"\nTop");
        pagedSource.replace("Name =\"PageNumber\"\nTop", "Name =\"PageNumber\"\nControlSource =\"=CStr([Page]) & Chr(47) & CStr([Pages])\"\nTop");
        AccessFormPreview paged("paged report",pagedSource,{},db);
        auto* reportPages=paged.findChild<QPrintPreviewWidget*>("access-report-pages");
        auto* reportTabs=paged.findChild<QTabWidget*>("access-report-view-tabs");
        auto* pageNumber=paged.findChild<QSpinBox*>("access-report-page-number");
        assert(reportPages && reportPages->pageCount()==2 && reportTabs && reportTabs->currentIndex()==0 && pageNumber);
        pageNumber->setValue(2);
        assert(reportPages->currentPage()==2 && makePreviewFormObject(paged)->member("Page",{}).asNumber()==2);
        assert(makePreviewFormObject(paged)->member("Pages",{}).asNumber()==2);
        pageNumber->setValue(1);
        paged.show(); QApplication::setActiveWindow(&paged);
        auto* navigationSurface=reportPages->findChild<QGraphicsView*>(); assert(navigationSurface);
        navigationSurface->setFocus(); QApplication::processEvents();
        QTest::keyClick(navigationSurface,Qt::Key_F5);
        assert(pageNumber==QApplication::focusWidget() || pageNumber->isAncestorOf(QApplication::focusWidget()));
        auto* pageEditor=pageNumber->findChild<QLineEdit*>(); assert(pageEditor);
        QTest::keyClicks(pageEditor,"2"); assert(reportPages->currentPage()==1);
        QTest::keyClick(pageEditor,Qt::Key_Return); assert(reportPages->currentPage()==2);
        navigationSurface->setFocus(); QApplication::processEvents();
        QTest::keyClick(navigationSurface,Qt::Key_PageUp); assert(reportPages->currentPage()==1);
        QTest::keyClick(navigationSurface,Qt::Key_PageDown); assert(reportPages->currentPage()==2);
        QTest::keyClick(navigationSurface,Qt::Key_Up); assert(reportPages->currentPage()==1);
        QTest::keyClick(navigationSurface,Qt::Key_Down); assert(reportPages->currentPage()==2);
        QTest::keyClick(navigationSurface,Qt::Key_Z); QApplication::processEvents();
        assert(reportPages->zoomMode()==QPrintPreviewWidget::CustomZoom && qAbs(reportPages->zoomFactor()-1.0)<0.001);
        QTest::keyClick(navigationSurface,Qt::Key_Z); QApplication::processEvents();
        assert(reportPages->zoomMode()==QPrintPreviewWidget::FitInView);
        pageNumber->setValue(1); paged.hide();
        const auto document=output.filePath("records.pdf");
        assert(paged.exportReportPdf(document));
        QFile pdf(document); assert(pdf.open(QIODevice::ReadOnly));
        const auto bytes=pdf.readAll();
        const QRegularExpression pagePattern("/Type\\s*/Page(?:\\s|/|>)");
        auto pages=pagePattern.globalMatch(QString::fromLatin1(bytes));
        int pageCount=0; while(pages.hasNext()) { pages.next(); ++pageCount; }
        assert(pageCount==2);
        const auto reader=QStandardPaths::findExecutable("pdftotext");
        if (!reader.isEmpty()) {
            QProcess textReader; textReader.start(reader,{"-layout",document,"-"});
            assert(textReader.waitForFinished(10000) && textReader.exitCode()==0);
            const auto text=QString::fromUtf8(textReader.readAllStandardOutput())
                .normalized(QString::NormalizationForm_KC);
            assert(text.count("first")==1 && text.count("second")==1);
            assert(text.indexOf("first")<text.indexOf("second"));
            assert(text.contains("10") && text.contains("20") && text.contains("3"));
            assert(text.count("1/2")==1 && text.count("2/2")==1);
        }

        assert(paged.findChild<QLineEdit*>("access-control-Name")->text()=="first");
        auto* previewAction=paged.findChild<QAction*>("access-report-preview");
        assert(previewAction && paged.findChild<QAction*>("access-report-print"));
        bool previewChecked=false;
        QTimer::singleShot(0,[&] {
            auto* dialog=qobject_cast<QPrintPreviewDialog*>(QApplication::activeModalWidget());
            assert(dialog);
            auto* pages=dialog->findChild<QPrintPreviewWidget*>();
            assert(pages && pages->pageCount()==2);
            previewChecked=true; dialog->reject();
        });
        previewAction->trigger();
        assert(previewChecked && paged.findChild<QLineEdit*>("access-control-Name")->text()=="first");
        paged.show(); QApplication::setActiveWindow(&paged); reportPages->setFocus(); QApplication::processEvents();
        for(const auto modifiers:{Qt::ControlModifier,Qt::NoModifier}) {
            QApplication::setActiveWindow(&paged);
            auto* surface=reportPages->findChild<QGraphicsView*>(); assert(surface);
            surface->setFocus(); QApplication::processEvents();
            bool printChecked=false;
            QTimer::singleShot(0,&paged,[&] {
                auto* dialog=qobject_cast<QPrintDialog*>(QApplication::activeModalWidget());
                assert(dialog && !dialog->testOption(QAbstractPrintDialog::PrintPageRange));
                printChecked=true; dialog->reject();
            });
            QTest::keyClick(surface,Qt::Key_P,modifiers);
            assert(printChecked && paged.findChild<QLineEdit*>("access-control-Name")->text()=="first");
        }
        auto* setupAction=paged.findChild<QAction*>("access-report-page-setup"); assert(setupAction);
        QApplication::setActiveWindow(&paged);
        auto* setupSurface=reportPages->findChild<QGraphicsView*>(); assert(setupSurface);
        setupSurface->setFocus(); QApplication::processEvents();
        bool pageSetupChecked=false;
        QTimer::singleShot(0,&paged,[&] {
            auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());
            assert(dialog && dialog->objectName()=="access-report-page-setup-dialog");
            dialog->findChild<QComboBox*>("report-page-orientation")->setCurrentIndex(1);
            dialog->findChild<QDoubleSpinBox*>("report-page-left")->setValue(25);
            pageSetupChecked=true;
            dialog->findChild<QDialogButtonBox*>("report-page-buttons")->button(QDialogButtonBox::Ok)->click();
        });
        QTest::keyClick(setupSurface,Qt::Key_S); assert(pageSetupChecked);
        const auto landscapePdf=output.filePath("landscape.pdf"); assert(paged.exportReportPdf(landscapePdf));
        QFile landscapeFile(landscapePdf); assert(landscapeFile.open(QIODevice::ReadOnly));
        const auto landscapeBytes=QString::fromLatin1(landscapeFile.readAll());
        const QRegularExpression mediaBox(R"(/MediaBox\s*\[\s*0\s+0\s+([0-9.]+)\s+([0-9.]+)\s*\])");
        const auto box=mediaBox.match(landscapeBytes); assert(box.hasMatch() && box.captured(1).toDouble()>box.captured(2).toDouble());
        if(!reader.isEmpty()) {
            QProcess textReader; textReader.start(reader,{"-bbox",landscapePdf,"-"});
            assert(textReader.waitForFinished(10000) && textReader.exitCode()==0);
            const auto text=QString::fromUtf8(textReader.readAllStandardOutput()).normalized(QString::NormalizationForm_KC);
            const QRegularExpression firstPosition(R"PDF(<word xMin="([0-9.]+)"[^>]*>first</word>)PDF");
            const auto position=firstPosition.match(text);
            assert(position.hasMatch() && position.captured(1).toDouble()>70 && position.captured(1).toDouble()<85);
        }
        bool cancelledSetup=false;
        QTimer::singleShot(0,&paged,[&] {
            auto* dialog=qobject_cast<QDialog*>(QApplication::activeModalWidget());
            assert(dialog && dialog->objectName()=="access-report-page-setup-dialog");
            auto* orientation=dialog->findChild<QComboBox*>("report-page-orientation"); assert(orientation->currentIndex()==1);
            assert(dialog->findChild<QDoubleSpinBox*>("report-page-left")->value()==25);
            orientation->setCurrentIndex(0);
            cancelledSetup=true;
            dialog->findChild<QDialogButtonBox*>("report-page-buttons")->button(QDialogButtonBox::Cancel)->click();
        });
        setupAction->trigger(); assert(cancelledSetup);
        assert(paged.exportReportPdf(output.filePath("landscape-after-cancel.pdf")));
        QFile cancelledFile(output.filePath("landscape-after-cancel.pdf")); assert(cancelledFile.open(QIODevice::ReadOnly));
        const auto cancelledBox=mediaBox.match(QString::fromLatin1(cancelledFile.readAll()));
        assert(cancelledBox.hasMatch() && cancelledBox.captured(1).toDouble()>cancelledBox.captured(2).toDouble());
        assert(paged.findChild<QLineEdit*>("access-control-Name")->text()=="first");
        paged.hide();
        assert(paged.findChild<QLineEdit*>("access-control-Info")->text()=="10");
        assert(paged.findChild<QLineEdit*>("access-control-ReportSummary")->text()=="3");
        auto footerSource=pagedSource;
        footerSource.replace("=Sum([id])","=1 / ([id]-1)");
        AccessFormPreview finalFooter("final footer report",footerSource,{},db);
        assert(finalFooter.exportReportPdf(output.filePath("final-footer.pdf")));
        assert(finalFooter.findChild<QLineEdit*>("access-control-Name")->text()=="first");
        assert(finalFooter.close());
        auto compoundSource=pagedSource;
        compoundSource.replace("=Sum([id])","=Sum(IIf([id]=1,10,20)) + Count(IIf([id]=1,Null,1)) + Count(Null)");
        AccessFormPreview compound("compound report",compoundSource,{},db);
        assert(compound.findChild<QLineEdit*>("access-control-ReportSummary")->text()=="31");
        assert(compound.exportReportPdf(output.filePath("compound.pdf")));
        assert(compound.findChild<QLineEdit*>("access-control-Name")->text()=="first");
        assert(query.exec("INSERT INTO records VALUES(3,'third')")); query.finish();
        makePreviewFormObject(compound)->member("Requery",{});
        assert(compound.findChild<QLineEdit*>("access-control-ReportSummary")->text()=="52");
        assert(compound.findChild<QPrintPreviewWidget*>("access-report-pages")->pageCount()==3);
        assert(compound.exportReportPdf(output.filePath("compound.pdf")));
        assert(compound.findChild<QLineEdit*>("access-control-ReportSummary")->text()=="52");
        assert(compound.close());
        assert(query.exec("DELETE FROM records WHERE id=3")); query.finish();
        assert(query.exec("CREATE TABLE native_report_stats(id INTEGER PRIMARY KEY,value REAL,label TEXT)"));
        assert(query.exec("INSERT INTO native_report_stats VALUES(1,10,'A'),(2,20,'B'),(3,NULL,NULL),(4,30,'C')")); query.finish();
        QStringList statsLines{"Begin Report","Width =9000","RecordSource =\"SELECT * FROM native_report_stats ORDER BY id\"",
            "Begin Section","Height =450","Begin TextBox","Name =\"RowId\"","ControlSource =\"id\"","Width =900","Height =300","End","End",
            "Begin FormFooter","Height =4200"};
        const QStringList functions{"avg","min","max","sum","count","compound","first","last","firstcompound","lastcompound","var","stdev","varp","stdevp"};
        for(int index=0;index<functions.size();++index) {
            const auto& function=functions[index]; const auto top=QString::number(index*300);
            const auto expression=function=="compound" ? QString("Avg(IIf([value]>10,[value]*2,Null))")
                : function=="firstcompound" ? QString("First(IIf([id]=1,Null,[value]))")
                : function=="lastcompound" ? QString("Last(IIf([id]=5,Null,[label]))")
                : function+((function=="first" || function=="last") ? "([label])" : "([value])");
            statsLines<<"Begin Label"<<"Name =\"Label_"+function+"\""<<"Caption =\"STAT_"+function+"\""<<"Top ="+top<<"Width =2800"<<"Height =270"<<"End"
                <<"Begin TextBox"<<"Name =\"Stat_"+function+"\""<<"ControlSource =\"="+expression+"\""<<"Top ="+top<<"Left =3000"<<"Width =3000"<<"Height =270"<<"End";
        }
        statsLines<<"End"<<"End"; const auto statsSource=statsLines.join('\n').toUtf8();
        AccessFormPreview stats("numeric report aggregates",statsSource,{},db);
        const auto checkStats=[&] {
            QSqlQuery reference(db);
            assert(reference.exec("SELECT AVG(value),MIN(value),MAX(value),SUM(value),COUNT(value),AVG(CASE WHEN value>10 THEN value*2 END) FROM native_report_stats") && reference.next());
            for(int index=0;index<6;++index) {
                bool valid=false; const auto actual=stats.findChild<QLineEdit*>("access-control-Stat_"+functions[index])->text().toDouble(&valid);
                assert(valid && qAbs(actual-reference.value(index).toDouble())<1e-10);
            }
            reference.finish();
            for(const auto& function:{QString("first"),QString("last")}) {
                assert(reference.exec("SELECT label FROM native_report_stats ORDER BY id "+QString(function=="last" ? "DESC" : "ASC")+" LIMIT 1") && reference.next());
                assert(stats.findChild<QLineEdit*>("access-control-Stat_"+function)->text()==reference.value(0).toString()); reference.finish();
            }
            assert(reference.exec("SELECT COUNT(value),SUM(value),SUM(value*value) FROM native_report_stats") && reference.next());
            const double count=reference.value(0).toDouble(),sum=reference.value(1).toDouble();
            const double variance=(reference.value(2).toDouble()-sum*sum/count)/(count-1);
            assert(qAbs(stats.findChild<QLineEdit*>("access-control-Stat_var")->text().toDouble()-variance)<1e-8);
            assert(qAbs(stats.findChild<QLineEdit*>("access-control-Stat_stdev")->text().toDouble()-std::sqrt(variance))<1e-8);
            const double population=variance*(count-1)/count;
            assert(qAbs(stats.findChild<QLineEdit*>("access-control-Stat_varp")->text().toDouble()-population)<1e-8);
            assert(qAbs(stats.findChild<QLineEdit*>("access-control-Stat_stdevp")->text().toDouble()-std::sqrt(population))<1e-8);
            assert(stats.findChild<QLineEdit*>("access-control-Stat_firstcompound")->text().isEmpty());
        };
        checkStats(); const auto statsPdf=output.filePath("stats.pdf"); assert(stats.exportReportPdf(statsPdf));
        if(!reader.isEmpty()) {
            QProcess textReader; textReader.start(reader,{"-layout",statsPdf,"-"});
            assert(textReader.waitForFinished(10000) && textReader.exitCode()==0);
            const auto text=QString::fromUtf8(textReader.readAllStandardOutput()).normalized(QString::NormalizationForm_KC);
            assert(QRegularExpression("STAT_avg\\s+20\\b").match(text).hasMatch());
            assert(QRegularExpression("STAT_min\\s+10\\b").match(text).hasMatch());
            assert(QRegularExpression("STAT_max\\s+30\\b").match(text).hasMatch());
            assert(QRegularExpression("STAT_compound\\s+50\\b").match(text).hasMatch());
        }
        auto reversedSource=statsSource; reversedSource.replace("ORDER BY id","ORDER BY id DESC");
        AccessFormPreview reversed("reverse report selectors",reversedSource,{},db);
        assert(reversed.findChild<QLineEdit*>("access-control-Stat_first")->text()=="C");
        assert(reversed.findChild<QLineEdit*>("access-control-Stat_last")->text()=="A");
        assert(reversed.close());
        assert(query.exec("INSERT INTO native_report_stats VALUES(5,40,'D')")); query.finish();
        auto statsObject=makePreviewFormObject(stats); (void)statsObject->member("Requery",{}); checkStats();
        assert(stats.exportReportPdf(statsPdf));
        assert(stats.findChild<QLineEdit*>("access-control-Stat_lastcompound")->text().isEmpty());
        if(!reader.isEmpty()) {
            QProcess textReader; textReader.start(reader,{"-layout",statsPdf,"-"});
            assert(textReader.waitForFinished(10000) && textReader.exitCode()==0);
            const auto text=QString::fromUtf8(textReader.readAllStandardOutput()).normalized(QString::NormalizationForm_KC);
            assert(QRegularExpression("STAT_first\\s+A\\b").match(text).hasMatch());
            assert(QRegularExpression("STAT_last\\s+D\\b").match(text).hasMatch());
        }
        assert(query.exec("UPDATE native_report_stats SET label=NULL WHERE id IN (1,5)")); query.finish();
        (void)statsObject->member("Requery",{});
        for(const auto& function:{QString("first"),QString("last")}) {
            const auto control=makePreviewControlsObject(stats)->member("Item",{linaccess::VbaValue::string(("Stat_"+function).toStdString())});
            assert(control.object->member("Value",{}).kind==linaccess::VbaValue::Kind::Null);
        }
        QStringList typedLines{"Begin Report","Width =9000",
            "RecordSource =\"SELECT * FROM native_report_stats WHERE value IS NOT NULL ORDER BY id\"",
            "Begin Section","Height =300","End","Begin FormFooter","Height =1200"};
        const QStringList typedNames{"mincash","maxcash","mindate","maxdate"};
        for(int index=0;index<typedNames.size();++index) {
            const auto expression=QString(index%2 ? "Max(" : "Min(")
                +(index<2 ? "CCur([value]))" : "CDate([value]))");
            typedLines<<"Begin TextBox"<<"Name =\""+typedNames[index]+"\""<<"ControlSource =\"="+expression+"\""
                <<"Top ="+QString::number(index*300)<<"Width =4000"<<"Height =270"<<"End";
        }
        typedLines<<"End"<<"End";
        AccessFormPreview typedStats("typed report extrema",typedLines.join('\n').toUtf8(),{},db);
        for(int index=0;index<typedNames.size();++index) {
            const auto control=makePreviewControlsObject(typedStats)->member("Item",{linaccess::VbaValue::string(typedNames[index].toStdString())});
            const auto value=control.object->member("Value",{});
            const int expected=index%2 ? 40 : 10;
            if(index<2) assert(value.kind==linaccess::VbaValue::Kind::Currency && value.currencyUnits==expected*10000);
            else assert(value.kind==linaccess::VbaValue::Kind::Date && value.number==expected);
        }
        assert(typedStats.exportReportPdf(output.filePath("typed-extrema.pdf")));
        assert(typedStats.close());
        assert(query.exec("CREATE TABLE native_currency_sums(id INTEGER PRIMARY KEY,amount TEXT)"));
        assert(query.exec("INSERT INTO native_currency_sums VALUES(1,'922337203685477.5806'),(2,'0.0001')")); query.finish();
        const QByteArray currencySource=R"FORM(Begin Report
Width =9000
RecordSource ="native_currency_sums"
Begin Section
Height =300
End
Begin FormFooter
Height =300
Begin TextBox
Name ="ExactTotal"
ControlSource ="=Sum(CCur([amount]))"
Width =6000
Height =270
End
End
End
)FORM";
        AccessFormPreview currencyStats("exact Currency sum",currencySource,{},db);
        const auto currencyControl=makePreviewControlsObject(currencyStats)->member("Item",{linaccess::VbaValue::string("ExactTotal")});
        const auto exactTotal=currencyControl.object->member("Value",{});
        assert(exactTotal.kind==linaccess::VbaValue::Kind::Currency && exactTotal.currencyUnits==INT64_MAX);
        const auto exactPdf=output.filePath("exact-currency-sum.pdf"); assert(currencyStats.exportReportPdf(exactPdf));
        QFile exactFile(exactPdf); assert(exactFile.open(QIODevice::ReadOnly)); const auto exactBytes=exactFile.readAll(); exactFile.close();
        assert(query.exec("INSERT INTO native_currency_sums VALUES(3,'0.0001')")); query.finish();
        (void)makePreviewFormObject(currencyStats)->member("Requery",{});
        assert(currencyStats.findChild<QLineEdit*>("access-control-ExactTotal")->property("la_calculation_error").toInt()==6);
        assert(!currencyStats.exportReportPdf(exactPdf));
        assert(exactFile.open(QIODevice::ReadOnly) && exactFile.readAll()==exactBytes); exactFile.close();
        assert(currencyStats.close());
        assert(query.exec("DROP TABLE native_currency_sums")); query.finish();
        assert(query.exec("UPDATE native_report_stats SET value=CASE WHEN id=1 THEN 10 ELSE NULL END")); query.finish();
        (void)statsObject->member("Requery",{});
        for(const auto& name:{"Stat_var","Stat_stdev"}) {
            const auto control=makePreviewControlsObject(stats)->member("Item",{linaccess::VbaValue::string(name)});
            assert(control.object->member("Value",{}).kind==linaccess::VbaValue::Kind::Null);
        }
        for(const auto& name:{"Stat_varp","Stat_stdevp"}) {
            const auto control=makePreviewControlsObject(stats)->member("Item",{linaccess::VbaValue::string(name)});
            const auto value=control.object->member("Value",{});
            assert(value.kind==linaccess::VbaValue::Kind::Number && value.asNumber()==0);
        }
        assert(query.exec("UPDATE native_report_stats SET value=NULL")); query.finish();
        (void)statsObject->member("Requery",{});
        for(const auto& function:functions) {
            const auto text=stats.findChild<QLineEdit*>("access-control-Stat_"+function)->text();
            assert(function=="count" ? text=="0" : text.isEmpty());
        }
        assert(stats.exportReportPdf(output.filePath("null-stats.pdf")));
        auto textStatsSource=statsSource; textStatsSource.replace("=min([value])","=min([label])");
        AccessFormPreview textStats("unsupported text minimum",textStatsSource,{},db);
        assert(textStats.findChild<QLineEdit*>("access-control-Stat_min")->property("la_calculation_error").toInt()==3251);
        QFile protectedStats(statsPdf); assert(protectedStats.open(QIODevice::ReadOnly)); const auto numericPdf=protectedStats.readAll(); protectedStats.close();
        assert(!textStats.exportReportPdf(statsPdf));
        assert(protectedStats.open(QIODevice::ReadOnly) && protectedStats.readAll()==numericPdf);
        assert(textStats.close() && stats.close());
        assert(query.exec("DROP TABLE native_report_stats")); query.finish();
        assert(paged.close());
        auto invalidSource=pagedSource;
        invalidSource.replace("=[id] * 10","=[missing]");
        AccessFormPreview invalid("invalid expression report",invalidSource,{},db);
        assert(!invalid.exportReportPdf(document));
        assert(invalid.close());
        auto shrinkingSource=pagedSource; shrinkingSource.replace("Name =\"Name\"","Name =\"Name\"\nCanShrink =-1");
        AccessFormPreview shrinking("shrinking report",shrinkingSource,{},db);
        assert(shrinking.findChild<QTabWidget*>("access-report-view-tabs")->currentIndex()==1);
        assert(!shrinking.exportReportPdf(document));
        QFile unchanged(document); assert(unchanged.open(QIODevice::ReadOnly) && unchanged.readAll()==bytes);
        assert(shrinking.close());
        query.prepare("UPDATE records SET name=? WHERE id=1");
        const QString paragraphs="ALPHA paragraph\nOMEGA paragraph with wrapped words";
        query.addBindValue(paragraphs); assert(query.exec()); query.finish();
        auto multilineSource=pagedSource;
        multilineSource.replace("Name =\"Name\"\nControlSource", "Name =\"Name\"\nHeight =1800\nControlSource");
        multilineSource.replace("Top =400","Top =2400");
        AccessFormPreview multiline("multiline report",multilineSource,{},db);
        assert(multiline.findChild<QLineEdit*>("access-control-Name")->text()==paragraphs);
        const auto multilinePdf=output.filePath("multiline.pdf"); assert(multiline.exportReportPdf(multilinePdf));
        if (!reader.isEmpty()) {
            QProcess textReader; textReader.start(reader,{"-layout",multilinePdf,"-"});
            assert(textReader.waitForFinished(10000) && textReader.exitCode()==0);
            const auto text=QString::fromUtf8(textReader.readAllStandardOutput()).normalized(QString::NormalizationForm_KC);
            assert(text.count("ALPHA paragraph")==1 && text.count("OMEGA paragraph")==1);
            assert(text.indexOf("ALPHA paragraph")<text.indexOf("OMEGA paragraph"));
            const auto omega=text.indexOf("OMEGA"),wrapped=text.indexOf("with",omega);
            assert(wrapped>omega && text.mid(omega,wrapped-omega).contains('\n'));
        }
        assert(multiline.close());
        auto maskedSource=multilineSource;
        maskedSource.replace("Name =\"Name\"\nHeight", "Name =\"Name\"\nInputMask =\"Password\"\nHeight");
        AccessFormPreview masked("masked multiline report",maskedSource,{},db);
        const auto maskedPdf=output.filePath("masked-multiline.pdf"); assert(masked.exportReportPdf(maskedPdf));
        if (!reader.isEmpty()) {
            QProcess textReader; textReader.start(reader,{"-layout",maskedPdf,"-"});
            assert(textReader.waitForFinished(10000) && textReader.exitCode()==0);
            const auto text=textReader.readAllStandardOutput();
            assert(!text.contains("ALPHA") && !text.contains("OMEGA"));
        }
        assert(masked.close());
        assert(query.exec("UPDATE records SET name='first' WHERE id=1")); query.finish();
        QStringList alphaLines,omegaLines;
        for(int line=0;line<44;++line) { alphaLines.append("ALPHA "+QString::number(line)); omegaLines.append("OMEGA "+QString::number(line)); }
        alphaLines.append("ALPHA_END"); omegaLines.append("OMEGA_END");
        for(int id=1;id<=2;++id) {
            query.prepare("UPDATE records SET name=? WHERE id=?");
            query.addBindValue((id==1 ? alphaLines : omegaLines).join('\n')); query.addBindValue(id); assert(query.exec()); query.finish();
        }
        for(int setting=1;setting<=3;++setting) {
            assert(query.exec("UPDATE records SET name=CASE id WHEN 1 THEN 'FIRST_BREAK' ELSE 'LAST_BREAK' END")); query.finish();
            auto forcedSource=pagedSource;
            forcedSource.replace("Height =9000","Height =900\nForceNewPage ="+QByteArray::number(setting));
            const auto forcedLayout=linaccess::parseAccessFormLayout(QString::fromUtf8(forcedSource));
            assert(forcedLayout.detailForceNewPage==setting && forcedLayout.reportPrintLimitations.isEmpty());
            AccessFormPreview forced("forced detail pages",forcedSource,{},db);
            const int expected=setting==3 ? 4 : 3;
            assert(forced.findChild<QPrintPreviewWidget*>("access-report-pages")->pageCount()==expected);
            const auto forcedPdf=output.filePath("forced.pdf"); assert(forced.exportReportPdf(forcedPdf));
            if(!reader.isEmpty()) {
                QProcess textReader; textReader.start(reader,{"-layout",forcedPdf,"-"});
                assert(textReader.waitForFinished(10000) && textReader.exitCode()==0);
                const auto text=QString::fromUtf8(textReader.readAllStandardOutput());
                const auto pages=text.split(QChar('\f'));
                const int firstPage=setting==2 ? 0 : 1;
                assert(pages[firstPage].count("FIRST_BREAK")==1 && pages[firstPage+1].count("LAST_BREAK")==1);
                assert(text.count("FIRST_BREAK")==1 && text.count("LAST_BREAK")==1);
            }
            assert(forced.close());
        }
        for(const auto flags:{QPair<int,int>{2,0},{0,1},{2,1}}) {
            auto separatedSource=pagedSource;
            separatedSource.replace("Height =9000","Height =900");
            separatedSource.replace("Begin FormHeader\nHeight =600","Begin FormHeader\nForceNewPage ="+QByteArray::number(flags.first)+"\nHeight =600");
            separatedSource.replace("Begin FormFooter\nHeight =600","Begin FormFooter\nForceNewPage ="+QByteArray::number(flags.second)+"\nHeight =600");
            separatedSource.replace("=Sum([id])","=Chr(83) & Chr(70) & CStr(Sum([id]))");
            AccessFormPreview separated("separated report bands",separatedSource,{},db);
            const int expected=1+(flags.first ? 1 : 0)+(flags.second ? 1 : 0);
            assert(separated.findChild<QPrintPreviewWidget*>("access-report-pages")->pageCount()==expected);
            const auto separatedPdf=output.filePath("separated.pdf"); assert(separated.exportReportPdf(separatedPdf));
            if(!reader.isEmpty()) {
                QProcess textReader; textReader.start(reader,{"-layout",separatedPdf,"-"});
                assert(textReader.waitForFinished(10000) && textReader.exitCode()==0);
                const auto text=QString::fromUtf8(textReader.readAllStandardOutput()); const auto pages=text.split(QChar('\f'));
                const int recordsPage=flags.first ? 1 : 0;
                assert(pages[recordsPage].contains("FIRST_BREAK") && pages[recordsPage].contains("LAST_BREAK"));
                assert(pages[expected-1].contains("SF3") && text.count("SF3")==1);
            }
            assert(separated.close());
        }
        for(int id=1;id<=2;++id) {
            query.prepare("UPDATE records SET name=? WHERE id=?");
            query.addBindValue((id==1 ? alphaLines : omegaLines).join('\n')); query.addBindValue(id); assert(query.exec()); query.finish();
        }
        auto sectionSource=pagedSource;
        sectionSource.replace("Height =9000","Height =900\nName =\"DetailFixture\"");
        const QByteArray sectionCode=R"VBA(Private retainedSection As Access.Section
Sub Report_Open(Cancel As Integer)
Set retainedSection=Me.Section(acDetail)
If Not TypeOf retainedSection Is Section Then Err.Raise 5
If retainedSection.Height<>900 Then Err.Raise 5
If retainedSection.Name<>"DetailFixture" Then Err.Raise 5
retainedSection.ForceNewPage=3
Me.Section(acHeader).ForceNewPage=2
Me.Section(acFooter).ForceNewPage=1
retainedSection.BackColor=255
End Sub
Sub ResetSections()
retainedSection.ForceNewPage=0
Me.Section(acHeader).ForceNewPage=0
Me.Section(acFooter).ForceNewPage=0
retainedSection.BackColor=65280
End Sub
)VBA";
        AccessFormPreview sectionReport("native sections",sectionSource,sectionCode,db);
        if(const auto result=sectionReport.eventResult("OnOpen"); result && result->status!=linaccess::VbaExecutionStatus::Completed)
            std::cerr<<"Synthetic Section Open: error="<<result->errorNumber<<", line="<<result->line<<", "<<result->explanation<<'\n';
        assert(sectionReport.eventResult("OnOpen") && sectionReport.eventResult("OnOpen")->status==linaccess::VbaExecutionStatus::Completed);
        auto sectionForm=makePreviewFormObject(sectionReport);
        auto section=sectionForm->member("Section",{linaccess::VbaValue::integer(0)}).object;
        assert(section==sectionForm->member("Section",{linaccess::VbaValue::string("detailfixture")}).object);
        assert(section->member("Parent",{}).object==sectionForm && section->member("ForceNewPage",{}).asNumber()==3);
        assert(sectionReport.findChild<QPrintPreviewWidget*>("access-report-pages")->pageCount()==4);
        auto* sectionCanvas=sectionReport.findChild<QWidget*>("access-form-canvas");
        const int sectionY=qRound(linaccess::parseAccessFormLayout(QString::fromUtf8(sectionSource)).sectionOffsets.value("Section")/15.0)+10;
        assert(section->member("BackColor",{}).asNumber()==255 && sectionCanvas->grab().toImage().pixelColor(200,sectionY)==QColor(Qt::red));
        std::vector<linaccess::VbaValue> sectionArgs;
        assert(sectionReport.executeProcedure("ResetSections",sectionArgs).status==linaccess::VbaExecutionStatus::Completed);
        assert(section->member("ForceNewPage",{}).asNumber()==0);
        assert(sectionReport.findChild<QPrintPreviewWidget*>("access-report-pages")->pageCount()==1);
        assert(section->member("BackColor",{}).asNumber()==65280 && sectionCanvas->grab().toImage().pixelColor(200,sectionY)==QColor(Qt::green));
        bool invalidSectionColor=false;
        try { section->assign("BackColor",linaccess::VbaValue::longInteger(-1)); }
        catch(const linaccess::VbaRuntimeError& error) { invalidSectionColor=error.number==3251; }
        assert(invalidSectionColor && section->member("BackColor",{}).asNumber()==65280);
        bool invalidSectionSetting=false;
        try { section->assign("ForceNewPage",linaccess::VbaValue::integer(4)); }
        catch(const linaccess::VbaRuntimeError& error) { invalidSectionSetting=error.number==5; }
        assert(invalidSectionSetting && section->member("ForceNewPage",{}).asNumber()==0);
        sectionReport.show(); QApplication::setActiveWindow(&sectionReport);
        auto* sectionPages=sectionReport.findChild<QPrintPreviewWidget*>("access-report-pages");
        auto* sectionSurface=sectionPages->findChild<QGraphicsView*>(); assert(sectionSurface);
        sectionSurface->setFocus(); QApplication::processEvents();
        QTest::keyClick(sectionSurface,Qt::Key_C);
        assert(!sectionReport.isVisible() && sectionReport.close());
        bool closedSection=false;
        try { (void)section->member("Height",{}); }
        catch(const linaccess::VbaRuntimeError& error) { closedSection=error.number==2450; }
        assert(closedSection);
        auto shortSource=pagedSource; shortSource.replace("Height =9000","Height =900");
        AccessFormPreview fixedGrowth("fixed growth baseline",shortSource,{},db);
        const auto originalName=fixedGrowth.findChild<QLineEdit*>("access-control-Name")->geometry();
        const auto originalInfo=fixedGrowth.findChild<QLineEdit*>("access-control-Info")->geometry();
        const auto originalCanvas=fixedGrowth.findChild<QWidget*>("access-form-canvas")->size();
        assert(fixedGrowth.findChild<QPrintPreviewWidget*>("access-report-pages")->pageCount()==1);
        assert(fixedGrowth.close());
        shortSource.replace("Name =\"Name\"","Name =\"Name\"\nCanGrow =-1");
        AccessFormPreview flowed("flowed report",shortSource,{},db);
        auto* flowedPages=flowed.findChild<QPrintPreviewWidget*>("access-report-pages");
        auto* flowedName=flowed.findChild<QLineEdit*>("access-control-Name");
        if(flowedPages->pageCount()!=2) std::cerr<<"CanGrow preview: pages="<<flowedPages->pageCount()<<", width="<<flowedName->width()<<", height="<<flowedName->height()<<", preferred="<<flowedName->sizeHint().height()<<", raw-breaks="<<flowedName->text().count('\n')<<", display-breaks="<<flowedName->displayText().count('\n')<<'\n';
        assert(flowedPages->pageCount()==2);
        assert(makePreviewControlsObject(flowed)->member("Item",{linaccess::VbaValue::string("Name")}).object->member("CanGrow",{}).truthy());
        const auto flowPdf=output.filePath("flowed.pdf"); assert(flowed.exportReportPdf(flowPdf));
        assert(flowed.findChild<QLineEdit*>("access-control-Name")->geometry()==originalName);
        assert(flowed.findChild<QLineEdit*>("access-control-Info")->geometry()==originalInfo);
        assert(flowed.findChild<QWidget*>("access-form-canvas")->size()==originalCanvas);
        assert(flowed.findChild<QLineEdit*>("access-control-Name")->text()==alphaLines.join('\n'));
        if (!reader.isEmpty()) {
            QProcess textReader; textReader.start(reader,{"-layout",flowPdf,"-"});
            assert(textReader.waitForFinished(10000) && textReader.exitCode()==0);
            const auto text=QString::fromUtf8(textReader.readAllStandardOutput()).normalized(QString::NormalizationForm_KC);
            const auto pages=text.split(QChar('\f')); assert(pages.size()>=2);
            assert(pages[0].count("ALPHA_END")==1 && pages[1].count("OMEGA_END")==1);
            assert(QRegularExpression("\\b10\\b").match(pages[0],pages[0].indexOf("ALPHA_END")+9).hasMatch());
            assert(QRegularExpression("\\b20\\b").match(pages[1],pages[1].indexOf("OMEGA_END")+9).hasMatch());
        }
        assert(flowed.close());
        auto parallelSource=shortSource;
        parallelSource.replace("Name =\"Info\"","Name =\"Info\"\nCanGrow =-1\nLeft =2000");
        parallelSource.replace("ControlSource =\"=[id] * 10\"","ControlSource =\"[name]\"");
        parallelSource.replace("Top =400","Top =0");
        AccessFormPreview parallel("parallel growth report",parallelSource,{},db);
        assert(parallel.findChild<QPrintPreviewWidget*>("access-report-pages")->pageCount()==2);
        const auto parallelPdf=output.filePath("parallel.pdf"); assert(parallel.exportReportPdf(parallelPdf));
        if (!reader.isEmpty()) {
            QProcess textReader; textReader.start(reader,{"-layout",parallelPdf,"-"});
            assert(textReader.waitForFinished(10000) && textReader.exitCode()==0);
            const auto text=QString::fromUtf8(textReader.readAllStandardOutput()).normalized(QString::NormalizationForm_KC);
            assert(text.count("ALPHA_END")==2 && text.count("OMEGA_END")==2);
        }
        assert(parallel.close());
        for(int id=1;id<=2;++id) {
            QStringList lines; for(int line=0;line<19;++line) lines.append((id==1 ? "ALPHA " : "OMEGA ")+QString::number(line));
            lines.append(id==1 ? "ALPHA_END" : "OMEGA_END");
            query.prepare("UPDATE records SET name=? WHERE id=?"); query.addBindValue(lines.join('\n')); query.addBindValue(id); assert(query.exec()); query.finish();
        }
        auto stackedSource=shortSource;
        stackedSource.replace("Name =\"Info\"","Name =\"Info\"\nCanGrow =-1");
        stackedSource.replace("ControlSource =\"=[id] * 10\"","ControlSource =\"[name]\"");
        AccessFormPreview stacked("stacked growth report",stackedSource,{},db);
        assert(stacked.findChild<QPrintPreviewWidget*>("access-report-pages")->pageCount()==2);
        const auto stackedPdf=output.filePath("stacked.pdf"); assert(stacked.exportReportPdf(stackedPdf));
        if (!reader.isEmpty()) {
            QProcess textReader; textReader.start(reader,{"-layout",stackedPdf,"-"});
            assert(textReader.waitForFinished(10000) && textReader.exitCode()==0);
            const auto text=QString::fromUtf8(textReader.readAllStandardOutput()).normalized(QString::NormalizationForm_KC);
            const auto pages=text.split(QChar('\f')); assert(pages.size()>=2);
            assert(pages[0].count("ALPHA_END")==2 && pages[1].count("OMEGA_END")==2);
        }
        assert(stacked.findChild<QLineEdit*>("access-control-Name")->geometry()==originalName);
        assert(stacked.findChild<QLineEdit*>("access-control-Info")->geometry()==originalInfo);
        assert(stacked.close());
        // One actual row spans multiple pages; every line must appear once.
        QStringList splitLines;
        for(int line=0;line<125;++line) splitLines.append("SPLIT_"+QString::number(line).rightJustified(3,'0'));
        for(const auto separator:{QChar('\n'),QChar(' ')}) {
            query.prepare("UPDATE records SET name=? WHERE id=1"); query.addBindValue(splitLines.join(separator)); assert(query.exec()); query.finish();
            auto splitSource=shortSource;
            splitSource.replace("RecordSource =\"records\"","RecordSource =\"SELECT * FROM records WHERE id=1\"");
            if(separator==' ') splitSource.replace("Name =\"Name\"","Name =\"Name\"\nFontUnderline =-1");
            AccessFormPreview splitReport("split detail report",splitSource,{},db);
            auto* splitPages=splitReport.findChild<QPrintPreviewWidget*>("access-report-pages");
            assert(splitPages->pageCount()>=2);
            const auto splitPdf=output.filePath("split.pdf"); assert(splitReport.exportReportPdf(splitPdf));
            if(!reader.isEmpty()) {
                QProcess textReader; textReader.start(reader,{"-layout",splitPdf,"-"});
                assert(textReader.waitForFinished(10000) && textReader.exitCode()==0);
                const auto text=QString::fromUtf8(textReader.readAllStandardOutput()).normalized(QString::NormalizationForm_KC);
                int last=-1;
                for(const auto& line:splitLines) { assert(text.count(line)==1); const int at=text.indexOf(line); assert(at>last); last=at; }
                const auto pages=text.split(QChar('\f'));
                assert(pages[0].contains("SPLIT_000") && !pages[0].contains("SPLIT_124"));
                assert(text.contains("1/"+QString::number(splitPages->pageCount())));
            }
            assert(splitReport.findChild<QLineEdit*>("access-control-Name")->geometry()==originalName);
            assert(splitReport.findChild<QLineEdit*>("access-control-Info")->geometry()==originalInfo);
            assert(splitReport.findChild<QWidget*>("access-form-canvas")->size()==originalCanvas);
            assert(splitReport.close());
        }
        QFile protectedFlow(flowPdf); assert(protectedFlow.open(QIODevice::ReadOnly)); const auto completeFlow=protectedFlow.readAll(); protectedFlow.close();
        query.prepare("UPDATE records SET name=? WHERE id=1"); query.addBindValue(QString("LONG\n").repeated(400)); assert(query.exec()); query.finish();
        AccessFormPreview tooLong("too long growing report",shortSource,{},db);
        assert(!tooLong.exportReportPdf(flowPdf));
        assert(tooLong.findChild<QLineEdit*>("access-control-Name")->geometry()==originalName);
        assert(tooLong.findChild<QLineEdit*>("access-control-Info")->geometry()==originalInfo);
        assert(protectedFlow.open(QIODevice::ReadOnly) && protectedFlow.readAll()==completeFlow);
        assert(tooLong.close());
        assert(query.exec("UPDATE records SET name=CASE id WHEN 1 THEN 'first' ELSE 'second' END")); query.finish();
        assert(banded.close());
    }
    AccessFormPreview report("report",source,code,db);
    assert(!report.cancelledByOpen() && report.hasDataSource());
    assert(report.eventResult("OnOpen") && report.eventResult("OnOpen")->status==linaccess::VbaExecutionStatus::Completed);
    assert(!report.eventResult("OnNoData"));
    auto* name = report.findChild<QLineEdit*>("access-control-Name");
    auto* info = report.findChild<QLineEdit*>("access-control-Info");
    auto* save = report.findChild<QPushButton*>("access-save-record");
    assert(name && name->text()=="first" && name->isReadOnly());
    assert(info && info->text()=="1:0"); assert(save && !save->isEnabled());
    assert(trace()==QStringList{"open"});
    auto reports = makePreviewReportsObject(report);
    assert(reports->member("Count",{}).asNumber()==1);
    assert(makePreviewFormsObject(report)->member("Count",{}).asNumber()==0);
    assert(reports->item("report").object->member("Name",{}).asText()=="report");
    bool readOnly = false;
    try { reports->item("report").object->assignItem("Name",linaccess::VbaValue::string("forbidden")); }
    catch(const linaccess::VbaRuntimeError& error) { readOnly=error.number==2448; }
    assert(readOnly && name->text()=="first");
    assert(query.exec("SELECT name FROM records WHERE id=1") && query.next() && query.value(0).toString()=="first"); query.finish();
    bool missing = false;
    try { (void)reports->item("missing"); } catch(const linaccess::VbaRuntimeError& error) { missing=error.number==2451; }
    assert(missing);
    assert(report.close());
    assert(report.eventResult("OnClose") && report.eventResult("OnClose")->status==linaccess::VbaExecutionStatus::Completed);
    assert(trace()==(QStringList{"open","close"}));
    clear();
    auto tableSource = source; tableSource.replace("RecordSource =\"ranked\"","RecordSource =\"records\"");
    AccessFormPreview tableReport("table report",tableSource,code,db);
    assert(tableReport.hasDataSource());
    assert(tableReport.findChild<QLineEdit*>("access-control-Name")->isReadOnly());
    assert(!tableReport.findChild<QPushButton*>("access-save-record")->isEnabled());
    assert(tableReport.close()); clear();
    auto cancelCode = code;
    cancelCode.replace("Sub Report_Open(Cancel As Integer)","Sub Report_Open(Cancel As Integer)\nCancel = -1");
    AccessFormPreview cancelled("cancelled",source,cancelCode,db);
    assert(cancelled.cancelledByOpen() && !cancelled.hasDataSource());
    assert(cancelled.eventResult("OnOpen")->cancelled && !cancelled.eventResult("OnNoData"));
    assert(trace()==QStringList{"open"}); clear();
    assert(query.exec("DELETE FROM records")); query.finish();
    AccessFormPreview empty("empty",source,code,db);
    assert(empty.hasDataSource() && !empty.cancelledByOpen());
    assert(empty.eventResult("OnNoData") && empty.eventResult("OnNoData")->status==linaccess::VbaExecutionStatus::Completed);
    assert(trace()==(QStringList{"open","nodata"}));
    assert(empty.close()); clear();
    auto noDataCancel = code; noDataCancel.replace("Cancel = 0","Cancel = -1");
    AccessFormPreview noRows("no rows",source,noDataCancel,db);
    assert(noRows.cancelledByOpen() && noRows.eventResult("OnNoData")->cancelled);
    assert(trace()==(QStringList{"open","nodata"})); clear();
    auto unavailable = source; unavailable.replace("RecordSource =\"ranked\"","RecordSource =\"missing\"");
    AccessFormPreview missingSource("missing source",unavailable,code,db);
    assert(!missingSource.hasDataSource() && !missingSource.eventResult("OnNoData"));
    assert(trace()==QStringList{"open"}); assert(missingSource.close());
}

void testOpenReport() {
    using linaccess::VbaValue;
    auto db=QSqlDatabase::addDatabase("QSQLITE","open-report-command"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery sql(db);
    assert(sql.exec("CREATE TABLE records(id INTEGER PRIMARY KEY,name TEXT)"));
    assert(sql.exec("INSERT INTO records VALUES(1,'first'),(2,'second')"));
    assert(sql.exec("CREATE TABLE events(value TEXT)"));
    assert(sql.exec("CREATE TABLE _la_access_artifacts(id INTEGER PRIMARY KEY,kind TEXT,name TEXT,source_bytes BLOB)"));
    const QByteArray layout=R"REPORT(Begin Report
RecordSource ="records"
OnOpen ="[Event Procedure]"
OnNoData ="[Event Procedure]"
OnClose ="[Event Procedure]"
Begin Section
Begin TextBox
Name ="value"
ControlSource ="name"
End
End
End
)REPORT";
    const QByteArray code=R"VBA(Sub Report_Open(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events VALUES('open')"
If Me.Tag="cancel" Then Cancel=1
End Sub
Sub Report_NoData(Cancel As Integer)
CurrentDb.Execute "INSERT INTO events VALUES('nodata')"
Cancel=1
End Sub
Sub Report_Close()
CurrentDb.Execute "INSERT INTO events VALUES('close')"
End Sub
)VBA";
    const auto artifact=[&](const QString& kind,const QString& name,const QByteArray& bytes) {
        sql.prepare("INSERT INTO _la_access_artifacts(kind,name,source_bytes) VALUES(?,?,?)"); sql.addBindValue(kind); sql.addBindValue(name); sql.addBindValue(bytes); assert(sql.exec()); sql.finish();
    };
    for (const auto* name:{"summary","empty","cancelled"}) {
        auto definition=layout; if(QString::fromLatin1(name)=="cancelled") definition.replace("Begin Report\n","Begin Report\nTag =\"cancel\"\n");
        artifact("report",name,definition); artifact("vba","Report_"+QString::fromLatin1(name),code);
    }
    auto missingSource=layout; missingSource.replace("RecordSource =\"records\"","RecordSource =\"missing_table\"");
    artifact("report","missing_source",missingSource); artifact("vba","Report_missing_source",code);
    artifact("report","failed",layout); artifact("vba","Report_failed","Sub Report_Open(Cancel As Integer)\nErr.Raise 17\nEnd Sub\n");
    artifact("report","kind_mismatch","Begin Form\nEnd\n");
    const QByteArray controller=R"VBA(Sub ShowReport()
DoCmd.OpenReport "summary",acViewPreview,,"[id]=2"
Me.Tag="caller retained"
End Sub
Sub ShowAgain()
Application.DoCmd.OpenReport "SUMMARY",acViewPreview,,"[id]=2",acWindowNormal
End Sub
)VBA";
    QMainWindow project;
    AccessFormPreview caller("controller","Begin Form\nEnd\n",controller,db,&project); caller.show(); QApplication::setActiveWindow(&caller); QApplication::processEvents();
    auto callerRoot=makePreviewFormObject(caller); auto commands=makePreviewApplicationObject(caller)->member("DoCmd",{}).object;
    const auto eventCount=[&](const char* event) { sql.prepare("SELECT COUNT(*) FROM events WHERE value=?"); sql.addBindValue(event); assert(sql.exec() && sql.next()); const int n=sql.value(0).toInt(); sql.finish(); return n; };
    std::vector<VbaValue> none; assert(caller.executeProcedure("ShowReport",none).status==linaccess::VbaExecutionStatus::Completed);
    assert(callerRoot->member("Tag",{}).asText()=="caller retained" && eventCount("open")==1);
    assert(makePreviewReportsObject(caller)->member("Count",{}).asNumber()==1);
    AccessFormPreview* report=nullptr;
    for(auto* widget:QApplication::topLevelWidgets()) if(auto* candidate=dynamic_cast<AccessFormPreview*>(widget);candidate && candidate!=&caller && candidate->isVisible() && makePreviewFormObject(*candidate)->member("Name",{}).asText()=="summary") { assert(!report); report=candidate; }
    assert(report && report->parentWidget()==&project && report->findChild<QToolBar*>("access-print-preview-toolbar"));
    auto reportRoot=makePreviewFormObject(*report); auto* field=report->findChild<QLineEdit*>("access-control-value"); assert(field && field->text()=="second" && field->isReadOnly());
    assert(reportRoot->member("Filter",{}).asText()=="[id]=2");
    assert(caller.executeProcedure("ShowAgain",none).status==linaccess::VbaExecutionStatus::Completed && eventCount("open")==1);
    const auto invokeError=[&](std::vector<VbaValue> args,int expected) {
        int error=0; try { commands->member("OpenReport",args); } catch(const linaccess::VbaRuntimeError& failure) { error=failure.number; }
        if(error!=expected) std::cerr<<"Report error contract: expected="<<expected<<", actual="<<error<<'\n';
        assert(error==expected);
    };
    invokeError({},450); invokeError({VbaValue{}},94);
    invokeError({VbaValue::string("summary")},3251);
    invokeError({VbaValue::string("summary"),VbaValue::numeric(1)},3251);
    invokeError({VbaValue::string("summary"),VbaValue::numeric(2147483648.0)},6);
    invokeError({VbaValue::string("summary"),VbaValue::numeric(2),VbaValue::string("filter")},3251);
    commands->member("OpenReport",{VbaValue::string("summary"),VbaValue::numeric(2),VbaValue::empty(),VbaValue::empty(),VbaValue::numeric(1)});
    assert(report->isHidden() && eventCount("open")==1);
    commands->member("OpenReport",{VbaValue::string("summary"),VbaValue::numeric(2),VbaValue::empty(),VbaValue::empty(),VbaValue::numeric(0)});
    assert(report->isVisible() && eventCount("open")==1);
    invokeError({VbaValue::string("missing"),VbaValue::numeric(2)},2103);
    assert(eventCount("open")==1);
    invokeError({VbaValue::string("cancelled"),VbaValue::numeric(2)},2501);
    invokeError({VbaValue::string("empty"),VbaValue::numeric(2),VbaValue::empty(),VbaValue::string("[id]=999")},2501);
    assert(eventCount("nodata")==1 && makePreviewReportsObject(caller)->member("Count",{}).asNumber()==1);
    invokeError({VbaValue::string("missing_source"),VbaValue::numeric(2)},3078);
    invokeError({VbaValue::string("failed"),VbaValue::numeric(2)},17);
    invokeError({VbaValue::string("kind_mismatch"),VbaValue::numeric(2)},2103);
    assert(eventCount("nodata")==1 && makePreviewReportsObject(caller)->member("Count",{}).asNumber()==1);

    assert(caller.close() && report->isVisible() && field->text()=="second");
    auto reportCommands=makePreviewApplicationObject(*report)->member("DoCmd",{}).object;
    reportCommands->member("Close",{VbaValue::numeric(3),VbaValue::string("summary")});
    assert(eventCount("close")==1);
    int closed=0; try { reportRoot->member("Name",{}); } catch(const linaccess::VbaRuntimeError& error) { closed=error.number; } assert(closed==2450);
    QApplication::processEvents();
}

void testOriginalOpenReport(const QDir& directory,const QJsonArray& objects,QSqlDatabase db) {
    const QString moduleHash="6e8e0a0e41e1a1768cff4a0b79a07205c7b0f799eb8b465af5397843ce048ed8";
    QByteArray code,layout; QString formName;
    const auto read=[&](const QJsonObject& artifact) {
        QFile file(directory.filePath(artifact["file"].toString())); assert(file.open(QIODevice::ReadOnly)); const auto bytes=file.readAll();
        assert(QCryptographicHash::hash(bytes,QCryptographicHash::Sha256).toHex()==artifact["sha256"].toString().toLatin1()); return bytes;
    };
    for(const auto& entry:objects) { const auto item=entry.toObject(); if(item["kind"]=="vba" && item["sha256"].toString()==moduleHash) { code=read(item); formName=item["name"].toString().mid(5); } }
    assert(!code.isEmpty());
    for(const auto& entry:objects) { const auto item=entry.toObject(); if(item["kind"]=="form" && item["name"].toString()==formName) layout=read(item); }
    assert(!layout.isEmpty());
    const QRegularExpression procedures("^\\s*(?:(?:Private|Public)\\s+)?Sub\\s+([\\p{L}_][\\p{L}\\p{N}_]*)\\s*\\([^\\n]*\\)([\\s\\S]*?)^\\s*End\\s+Sub",QRegularExpression::MultilineOption|QRegularExpression::CaseInsensitiveOption);
    const QRegularExpression preview("DoCmd\\s*\\.\\s*OpenReport\\s+([^,\\r\\n]+),\\s*acPreview\\b",QRegularExpression::CaseInsensitiveOption);
    QString procedure,target; auto matches=procedures.globalMatch(linaccess::decodeVbaSource(code));
    while(matches.hasNext()) {
        const auto candidate=matches.next(); const auto call=preview.match(candidate.captured(2)); if(!call.hasMatch()) continue;
        assert(procedure.isEmpty()); procedure=candidate.captured(1); const auto argument=call.captured(1).trimmed();
        if(argument.startsWith('"') && argument.endsWith('"')) target=argument.mid(1,argument.size()-2);
        else { const QRegularExpression value("\\b"+QRegularExpression::escape(argument)+"\\s*=\\s*\"([^\"]+)\"",QRegularExpression::CaseInsensitiveOption); const auto literal=value.match(candidate.captured(2)); assert(literal.hasMatch()); target=literal.captured(1); }
    }
    assert(!procedure.isEmpty() && !target.isEmpty());
    QByteArray reportBytes;
    for(const auto& entry:objects) { const auto item=entry.toObject(); if(item["kind"]=="report" && item["name"].toString()==target) {
        assert(item["sha256"].toString()=="30f1ebe8559d44c1ca9f965948b55e556460d560286f669cac8f43a298741a6f"); reportBytes=read(item);
    } }
    assert(!reportBytes.isEmpty()); const auto reportLayout=linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(reportBytes));
    const auto callerLayout=linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(layout)); QString buttonName;
    for(const auto& control:callerLayout.controls) {
        const auto prefix=control.eventProcPrefix.isEmpty() ? control.name : control.eventProcPrefix;
        if(control.type=="CommandButton" && control.onClick.compare("[Event Procedure]",Qt::CaseInsensitive)==0 && (prefix+"_Click").compare(procedure,Qt::CaseInsensitive)==0) buttonName=control.name;
    }
    assert(!buttonName.isEmpty()); QSqlQuery guard(db); assert(guard.exec("SAVEPOINT original_open_report"));
    QMainWindow project; auto* toolbar=project.addToolBar("Project"); toolbar->addAction("Test command"); project.menuBar()->addMenu("Project")->addAction(toolbar->actions().first());
    AccessFormPreview caller(formName,layout,code,db,&project); caller.show(); QApplication::setActiveWindow(&caller); QApplication::processEvents();
    const auto verify=[&] {
        assert(makePreviewReportsObject(caller)->member("Count",{}).asNumber()==1);
        AccessFormPreview* opened=nullptr;
        for(auto* widget:QApplication::topLevelWidgets()) if(auto* candidate=dynamic_cast<AccessFormPreview*>(widget); candidate && candidate!=&caller && candidate->isVisible() && makePreviewFormObject(*candidate)->member("Name",{}).asText()==target.toStdString()) { assert(!opened); opened=candidate; }
        assert(opened && opened->hasDataSource() && opened->findChild<QToolBar*>("access-print-preview-toolbar"));
        if(const auto event=opened->eventResult("OnOpen")) assert(event->status==linaccess::VbaExecutionStatus::Completed);
        QString querySource;
        if(db.tables().contains(reportLayout.recordSource,Qt::CaseInsensitive)) { auto table=reportLayout.recordSource; table.replace('"',"\"\""); querySource="SELECT * FROM \""+table+"\""; }
        else if(reportLayout.recordSource.startsWith("SELECT",Qt::CaseInsensitive) || reportLayout.recordSource.startsWith("PARAMETERS",Qt::CaseInsensitive)) querySource=reportLayout.recordSource;
        else querySource=linaccess::accessQuerySource(db,reportLayout.recordSource);
        const auto plan=linaccess::compileAccessQuery(querySource); QSqlQuery reference(db); assert(reference.exec(plan.sql) && reference.next());
        QVector<QString> referenceTypes;(void)linaccess::accessQueryColumns(db,plan,&referenceTypes);
        const auto columns=reference.record(); int cells=0;
        for(const auto& control:reportLayout.controls) {
            if(control.type!="TextBox") continue; auto fieldName=control.controlSource.trimmed();
            if(fieldName.startsWith('[') && fieldName.endsWith(']')) fieldName=fieldName.mid(1,fieldName.size()-2);
            const auto index=columns.indexOf(fieldName); if(index<0) continue;
            auto* field=opened->findChild<QLineEdit*>("access-control-"+control.name); assert(field && field->isReadOnly());
            const auto value=reference.value(index); verifyOriginalBoundText(*opened,control,field,value,referenceTypes.value(index)); if(!value.isNull()) ++cells;
        }
        assert(cells>0); reference.finish();
        makePreviewApplicationObject(caller)->member("DoCmd",{}).object->member("Close",{linaccess::VbaValue::numeric(3),linaccess::VbaValue::string(target.toStdString())});
    };
    std::vector<linaccess::VbaValue> none;
    const auto outcome=caller.executeProcedure(procedure,none);
    std::cerr<<"Original report preview: status="<<static_cast<int>(outcome.status)<<", error="<<outcome.errorNumber<<", line="<<outcome.line<<'\n';
    assert(outcome.status==linaccess::VbaExecutionStatus::Completed); verify();
    auto* button=caller.findChild<QPushButton*>("access-control-"+buttonName); assert(button);
    button->setEnabled(true); button->show(); button->click(); verify();
    const auto clicked=caller.eventResult("control:"+buttonName+":Click"); assert(clicked && clicked->status==linaccess::VbaExecutionStatus::Completed);
    assert(caller.close()); assert(guard.exec("ROLLBACK TO original_open_report") && guard.exec("RELEASE original_open_report"));
}
