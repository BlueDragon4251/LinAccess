#include "access_form_import.hpp"
#include "access_catalog.hpp"
#include "access_table_import.hpp"
#include "access_form_preview.hpp"
#include "access_form_objects.hpp"
#include "vba_source_codec.hpp"

#include <QApplication>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QLineEdit>
#include <QComboBox>
#include <QRegularExpression>
#include <QTemporaryDir>
#include <QMessageBox>
#include <QTimer>
#include <QPushButton>
#include <QSqlDatabase>
#include <QSqlQuery>

#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <iostream>

void testFormEvents();
void testEventPrefixes();
void testFilters();
void testMessageBoxes();
void testInputBoxes();
void testRecordsetNavigation();
void testRowSourceTypes();
void testRecordDelete();
void testRecordClipboard();
void testOriginalClipboard(const QDir&,const QJsonArray&,const QString&);
void testEval();
void testOriginalEval(const QDir&,const QJsonArray&);
void testOriginalRecordsetSearch(const QDir&,const QJsonArray&);
void testCalendarValues();
void testModuleStates(const QDir&,const QJsonArray&);
void testCloseCommands();
void testDocumentLifetimes();
void testDynamicCombos();
void testKeyboardCommands();
void testOpenFormCommands();
void testActiveDocumentCommands();
void testOptionGroups();
void testSendKeys();
void testCalculatedCaching();
void testComboDropdown();
void testOriginalMessageBox(const QDir&,const QJsonArray&);
void testSubforms();
void testFormObjects();
void testReportEvents();
void testReportCorpus(const QDir&,const QJsonArray&,const QString&);
void testOriginalControls(const QDir&,const QJsonArray&);
void testOriginalAsciiKeyPress(const QDir&,const QJsonArray&);

int main(int argc, char** argv) {
    QApplication app(argc, argv);
    assert(argc >= 2);
    const QDir directory(QString::fromLocal8Bit(argv[1]));
    QFile manifest(directory.filePath("manifest.json"));
    assert(manifest.open(QIODevice::ReadOnly));
    const auto objects = QJsonDocument::fromJson(manifest.readAll()).object().value("objects").toArray();
    const auto unknownBinary = linaccess::parseAccessFormLayout(R"FORM(Begin Form
UnknownProp = {305 ,452 ,1 ,1 ,1 } Begin
    0x12345678
End
Caption ="Begin"
Begin Section
Height =1000
Begin TextBox
Name ="Preserved"
End
End
End
)FORM");
    assert(unknownBinary.caption == "Begin");
    assert(unknownBinary.controls.size() == 1 && unknownBinary.controls[0].name == "Preserved");
    int scanned = 0;
    int controlCount = 0;
    int buttons = 0;
    int prefixes = 0;
    for (const auto& item : objects) {
        const auto object = item.toObject();
        if (object.value("kind") != "form") continue;
        QFile file(directory.filePath(object.value("file").toString()));
        assert(file.open(QIODevice::ReadOnly));
        const auto bytes = file.readAll();
        const auto form = linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(bytes));
        assert(form.width > 0);
        assert(form.height > 0);
        for (const auto& control : form.controls) {
            if (!control.eventProcPrefix.isEmpty()) ++prefixes;
            if (!control.visible) continue;
            if (control.type == "CommandButton") ++buttons;
        }
        controlCount += static_cast<int>(form.controls.size());
        ++scanned;
        {
            // The actual Qt widgets must instantiate and inherit geometry.
            auto db = QSqlDatabase::addDatabase("QSQLITE", "layout-test-" + QString::number(scanned));
            db.setDatabaseName(":memory:");
            assert(db.open());
            AccessFormPreview dialog(object.value("name").toString(), bytes, {}, db);
            assert(dialog.renderedControls() > 0);
            auto objectModel = makePreviewFormObject(dialog);
            auto controlsModel = objectModel->member("Controls",{}).object;
            const auto items = controlsModel->enumerate();
            assert(items.size() == controlsModel->member("Count",{}).asNumber());
            assert(items.size() <= static_cast<std::size_t>(dialog.renderedControls()));
            for (const auto& item : items) {
                assert(!item.object->member("Name",{}).asText().empty());
                assert(item.object->member("ControlType",{}).asNumber() >= 100);
                const auto name=QString::fromStdString(item.object->member("Name",{}).asText());
                for (const auto& definition:form.controls)
                    if (definition.name==name && !definition.eventProcPrefix.isEmpty())
                        assert(item.object->member("EventProcPrefix",{}).asText()==definition.eventProcPrefix.toStdString());
            }
            assert(dialog.findChild<QWidget*>("access-form-canvas") != nullptr);
            db.close();
        }
    }
    assert(scanned == 47);
    assert(controlCount > 150);
    assert(buttons > 0);
    assert(prefixes > 0);
    {
        // Exercise real Qt click -> imported VBA -> SQLite mutation, with no
        // synthetic confirmation dialog and no access to the original MDB.
        auto db = QSqlDatabase::addDatabase("QSQLITE", "layout-execution-test");
        db.setDatabaseName(":memory:");
        assert(db.open());
        QSqlQuery query(db);
        assert(query.exec("CREATE TABLE recipes (id INTEGER PRIMARY KEY, name TEXT)"));
        const QByteArray form = R"FORM(Begin Form
    Width =4200
    RecordSource ="recipes"
    Caption ="Keksverwaltung"
    OnOpen ="[Event Procedure]"
    Begin Section
        Height =2000
        Begin
            Begin TextBox
                Name ="Name"
                ControlSource ="name"
                Left =100
                Top =100
                Width =2000
                Height =300
            End
            Begin CommandButton
                Name ="btnInsert"
                OnClick ="[Event Procedure]"
                Caption ="Speichern"
                Left =100
                Top =600
                Width =1700
                Height =400
            End
        End
    End
End
)FORM";
        const QByteArray vba = R"VBA(Private Sub Form_Open(Cancel As Integer)
Me!Name = "Nico"
End Sub
Private Sub btnInsert_Click()
CurrentDb.Execute "INSERT INTO recipes (name) VALUES ('Keks')"
DoCmd.Requery
End Sub
)VBA";
        AccessFormPreview dialog("frmRezepte", form, vba, db);
        assert(dialog.renderedControls() == 2);
        auto* name = dialog.findChild<QLineEdit*>("access-control-Name");
        assert(name && name->text() == "Nico");
        auto* action = dialog.findChild<QPushButton*>("access-control-btnInsert");
        assert(action);
        action->click();
        assert(query.exec("SELECT name FROM recipes"));
        assert(query.next() && query.value(0).toString() == "Keks");
        db.close();
    }
    {
        auto db = QSqlDatabase::addDatabase("QSQLITE", "original-keypad-events");
        db.setDatabaseName(":memory:"); assert(db.open());
        QTimer acknowledge;
        QObject::connect(&acknowledge, &QTimer::timeout, [] {
            for (auto* widget : QApplication::topLevelWidgets())
                if (auto* message = qobject_cast<QMessageBox*>(widget)) message->accept();
        });
        acknowledge.start(10);
        int executedKeys = 0;
        for (const auto& value : objects) {
            const auto module = value.toObject();
            if (module["kind"] != "vba") continue;
            QFile file(directory.filePath(module["file"].toString())); assert(file.open(QIODevice::ReadOnly));
            const auto bytes = file.readAll();
            const auto code = linaccess::decodeVbaSource(bytes);
            if (!code.contains("Me!gegeben = Me!gegeben & C")) continue;
            const auto formName = module["name"].toString().mid(5);
            QByteArray form;
            for (const auto& objectValue : objects) {
                const auto object = objectValue.toObject();
                if (object["kind"] != "form" || object["name"] != formName) continue;
                QFile formFile(directory.filePath(object["file"].toString()));
                assert(formFile.open(QIODevice::ReadOnly)); form = formFile.readAll();
            }
            assert(!form.isEmpty());
            AccessFormPreview preview(formName, form, bytes, db);
            QLineEdit* given = nullptr;
            for (auto* editor : preview.findChildren<QLineEdit*>())
                if (editor->objectName().compare("access-control-gegeben", Qt::CaseInsensitive) == 0) given = editor;
            assert(given);
            const auto layout = linaccess::parseAccessFormLayout(linaccess::decodeVbaSource(form));
            for (const auto& control : layout.controls) {
                auto event = control.name;
                event.replace(QRegularExpression("[^\\p{L}\\p{N}_]"), "_");
                const QRegularExpression handler("(?im)^\\s*(?:Private\\s+|Public\\s+)?Sub\\s+"
                    + QRegularExpression::escape(event + "_Click")
                    + R"re(\(\)\s*\r?\n\s*KeyCapt\s+"([^"]*)"\s*\r?\n\s*End Sub)re");
                const auto match = handler.match(code);
                if (!match.hasMatch()) continue;
                auto* button = preview.findChild<QPushButton*>("access-control-" + control.name); assert(button);
                given->clear(); button->click();
                assert(given->text() == match.captured(1)); ++executedKeys;
            }
        }
        assert(executedKeys >= 10);
        std::cout << "Original keypad button/parameterized-handler executions: " << executedKeys << '\n';
    }
    testFormEvents();
    testEventPrefixes();
    testMessageBoxes();
    testInputBoxes();
    testRecordsetNavigation();
    testRowSourceTypes();
    testRecordDelete();
    testRecordClipboard();
    testEval();
    testOriginalEval(directory,objects);
    testOriginalRecordsetSearch(directory,objects);
    testCalendarValues();
    testModuleStates(directory,objects);
    testCloseCommands();
    testDocumentLifetimes();
    testDynamicCombos();
    testKeyboardCommands();
    testOriginalAsciiKeyPress(directory,objects);
    testOpenFormCommands();
    testActiveDocumentCommands();
    testOptionGroups();
    testSendKeys();
    testCalculatedCaching();
    testComboDropdown();
    testOriginalMessageBox(directory,objects);
    testFilters();
    testFormObjects();
    testSubforms();
    testReportEvents();
    testOriginalControls(directory,objects);
    if (argc >= 3) {
        testOriginalClipboard(directory,objects,QString::fromLocal8Bit(argv[2]));
        testReportCorpus(directory,objects,QString::fromLocal8Bit(argv[2]));
        QTemporaryDir temporary;
        assert(temporary.isValid());
        auto db = QSqlDatabase::addDatabase("QSQLITE", "original-kassa-dao-event");
        db.setDatabaseName(temporary.filePath("copy.linaccess"));
        assert(db.open());
        QSqlQuery query(db);
        assert(query.exec("CREATE TABLE _la_import_log(source_file,kind,name,status,details)"));
        assert(linaccess::importAccessDatabase(db, QString::fromLocal8Bit(argv[2])) == 29);
        QByteArray form, code;
        for (const auto& value : objects) {
            const auto object = value.toObject();
            if ((object.value("kind") == "form" && object.value("name") == "Tagesende oder Ende")
                || (object.value("kind") == "vba" && object.value("name") == "Form_Tagesende oder Ende")) {
                QFile file(directory.filePath(object.value("file").toString()));
                assert(file.open(QIODevice::ReadOnly));
                if (object.value("kind") == "form") form = file.readAll(); else code = file.readAll();
            }
        }
        assert(!form.isEmpty() && !code.isEmpty());
        // Extract only the source table identifier; execute the original whole
        // handler and original form definition, without changing their bytes.
        const auto source = linaccess::decodeVbaSource(code);
        const QRegularExpression tablePattern(R"re(SF_Ende_Kassa_Click\(\)[\s\S]*?OpenRecordset\("([^"]+)")re");
        const auto match = tablePattern.match(source);
        assert(match.hasMatch());
        auto table = match.captured(1); table.replace('"', "\"\"");
        const auto count = [&] {
            assert(query.exec("SELECT COUNT(*) FROM \"" + table + "\""));
            assert(query.next()); return query.value(0).toInt();
        };
        const int before = count();
        AccessFormPreview dialog("Tagesende oder Ende", form, code, db);
        for (const auto& name : {"BedienerTmp", "BedienerNameTmp"}) {
            auto* editor = dialog.findChild<QLineEdit*>("access-control-" + QString(name));
            assert(editor);
            editor->setText(QString(name) == "BedienerTmp" ? "1" : "LinAccess regression");
        }
        auto* button = dialog.findChild<QPushButton*>("access-control-SF Ende Kassa");
        assert(button);
        button->click();
        assert(count() == before + 1);
        assert(query.exec("SELECT PersonalNr,Bediener,Zeitpunkt FROM \"" + table + "\" ORDER BY rowid DESC LIMIT 1"));
        assert(query.next() && query.value(0).toInt() == 1);
        assert(query.value(1).toString() == "LinAccess regression");
        assert(!query.value(2).toString().isEmpty());
        std::cout << "Original Kassa form click executed DAO AddNew/Update on isolated database copy\n";
    }
    std::cout << "Access 2003 layout tests passed: " << scanned
              << " forms, " << controlCount << " controls, " << buttons << " declared-visible buttons, " << prefixes << " stored event prefixes\n";
}
