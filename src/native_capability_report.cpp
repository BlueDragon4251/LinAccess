#include "native_capability_report.hpp"
#include "execution_capabilities.hpp"
#include <QDialog>
#include <QMap>
#include <QDialogButtonBox>
#include <QVBoxLayout>
#include <QLabel>
#include <QTreeWidget>
#include <QPushButton>
#include <QJsonDocument>
#include <QFileDialog>
#include <QSaveFile>
#include <QMessageBox>
void showNativeCapabilityReport(QWidget* owner,QSqlDatabase db) {
    try {
        const auto report=linaccess::projectCapabilityReport(db);QDialog dialog(owner);dialog.setObjectName("project-capability-report");dialog.setWindowTitle("Kompatibilität und Ausführungsbeobachtungen");dialog.resize(980,600);auto* layout=new QVBoxLayout(&dialog);
        auto* explanation=new QLabel("Syntaxannahme ist kein Ausführungsnachweis. Beobachtungen betreffen tatsächlich versuchte Aufrufe dieser Sitzung und Quellrevision; gleichlautende Quellen beweisen keine Objektinstanz. Dieser Bericht startet keine Prozeduren.",&dialog);explanation->setWordWrap(true);layout->addWidget(explanation);
        auto* tree=new QTreeWidget(&dialog);tree->setObjectName("project-capability-items");tree->setHeaderLabels({"Objekt / Prozedur","Art","Einstufung","Ausführung / letzter Quelllauf"});layout->addWidget(tree);
        const auto label=[](const QJsonValue& value) {const auto key=value.toString();static const QMap<QString,QString> labels{{"native","Nativ"},{"imported","Importiert"},{"native_codebehind","Nativer Formular-/Berichtscode"},{"physical","Physische Tabelle"},{"module","Modul"},{"vba","VBA"},{"query","Abfrage"},{"form","Formular"},{"report","Bericht"},{"macro","Makro"},{"table","Tabelle"},{"not_observed","Kein Lauf beobachtet"},{"see_procedures","Siehe Prozeduren"},{"matching_source_observed","Passende Quelle beobachtet"},{"completed","Beendet"},{"stopped","Abgebrochen"},{"error","Fehler"},{"unsupported","Nicht unterstützt"},{"procedure_missing","Prozedur fehlt"},{"unclassified","Nicht eingestuft"},{"invalid_definition","Ungültige Definition"},{"source_unavailable","Quelle fehlt"},{"source_limit_exceeded","Quellgrenze überschritten"},{"procedure_inventory_only","Prozeduren inventarisiert"},{"syntax_accepted_runtime_unproven","Syntax angenommen; Lauf unbewiesen"},{"sql_compiled_runtime_unproven","SQL übersetzt; Lauf unbewiesen"},{"sql_compile_failed","SQL-Übersetzung fehlgeschlagen"},{"document_requires_live_context","Benötigt Dokumentkontext"},{"macro_requires_project_context","Benötigt Makro-Projektkontext"},{"macro_syntax_accepted_runtime_unproven","Makrosyntax angenommen; Lauf unbewiesen"},{"macro_definition_failed","Makrodefinition nicht unterstützt"},{"schema_present_data_unproven","Schema vorhanden; Datenverhalten unbewiesen"}};return labels.value(key,key);};
        for(const auto& value:report["artifacts"].toArray()) {const auto item=value.toObject();auto* root=new QTreeWidgetItem(tree,{item["name"].toString(),label(item["origin"])+" / "+label(item["kind"]),label(item["capability"]),label(item["execution"])+" / "+label(item["last_source_result"])});for(const auto& entry:item["procedures"].toArray()) {const auto procedure=entry.toObject();new QTreeWidgetItem(root,{procedure["name"].toString(),procedure["kind"].toString(),label(procedure["capability"]),label(procedure["execution"])+" / "+label(procedure["last_source_result"])});}}
        for(int column=0;column<4;++column) tree->resizeColumnToContents(column);
        auto* buttons=new QDialogButtonBox(QDialogButtonBox::Close,&dialog);auto* save=buttons->addButton("Bericht speichern…",QDialogButtonBox::ActionRole);layout->addWidget(buttons);QObject::connect(buttons,&QDialogButtonBox::rejected,&dialog,&QDialog::reject);
        QObject::connect(save,&QPushButton::clicked,&dialog,[&] {const auto path=QFileDialog::getSaveFileName(&dialog,"Bericht speichern",{},"JSON (*.json)");if(path.isEmpty()) return;QSaveFile file(path);const auto bytes=QJsonDocument(report).toJson();if(!file.open(QIODevice::WriteOnly) || file.write(bytes)!=bytes.size() || !file.commit()) QMessageBox::warning(&dialog,"Bericht speichern",file.errorString());});dialog.exec();
    }catch(const std::exception& failure) {QMessageBox::warning(owner,"Kompatibilitätsbericht",QString::fromUtf8(failure.what()));}
}
