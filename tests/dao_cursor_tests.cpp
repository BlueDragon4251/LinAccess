#include "dao_database.hpp"
#include <QCoreApplication>
#include <QSqlQuery>
#ifdef NDEBUG
#undef NDEBUG
#endif
#include <cassert>
#include <functional>

void testDaoCursors() {
    using namespace linaccess;
    auto db=QSqlDatabase::addDatabase("QSQLITE","dao-cursors"); db.setDatabaseName(":memory:"); assert(db.open());
    QSqlQuery query(db);
    assert(query.exec("CREATE TABLE cursor_values(id INTEGER PRIMARY KEY,label TEXT,score INTEGER)"));
    assert(query.exec("INSERT INTO cursor_values VALUES(1,'one',10),(2,'two',20),(3,NULL,20),(4,'four',30)"));
    const auto expect=[](int number,const std::function<void()>& action) {
        int caught=0; try { action(); } catch (const VbaRuntimeError& failure) { caught=failure.number; }
        assert(caught==number);
    };
    auto database=makeDaoDatabase(db);
    auto source=database->member("OpenRecordset",{VbaValue::string("cursor_values")}).object;
    auto clone=source->member("Clone",{}).object;
    assert(source->item("id").asNumber()==1 && clone->member("BOF",{}).truthy() && !clone->member("EOF",{}).truthy());
    expect(3021,[&] { (void)clone->item("id"); });
    const auto find=[](const std::shared_ptr<VbaObject>& rows,const char* method,const char* criteria) {
        rows->member(method,{VbaValue::string(criteria)});
    };
    find(clone,"FindFirst","label Is Null"); assert(clone->item("id").asNumber()==3);
    find(clone,"FindNext","score >= 20"); assert(clone->item("id").asNumber()==4);
    find(clone,"FindPrevious","score >= 20"); assert(clone->item("id").asNumber()==3);
    find(clone,"FindLast","score >= 20"); assert(clone->item("id").asNumber()==4 && source->item("id").asNumber()==1);
    find(clone,"FindFirst","id=999"); assert(clone->member("NoMatch",{}).truthy() && clone->member("EOF",{}).truthy());
    expect(3021,[&] { (void)clone->member("Bookmark",{}); });
    clone->member("MoveFirst",{}); assert(!clone->member("NoMatch",{}).truthy());
    find(clone,"FindFirst","label Like 't*'"); assert(clone->item("id").asNumber()==2);
    const auto second=clone->member("Bookmark",{}); assert(second.kind==VbaValue::Kind::ByteArray && !second.bytes.empty());
    source->assign("Bookmark",second); assert(source->item("id").asNumber()==2);
    const bool legacyDqs=query.exec("SELECT \"legacy_string\""); query.finish();
    expect(3070,[&] { find(clone,"FindFirst","[missing_field]=1"); });
    assert(clone->item("id").asNumber()==2 && !clone->member("NoMatch",{}).truthy());
    assert(query.exec("SELECT \"legacy_string\"")==legacyDqs); query.finish();
    expect(94,[&] { clone->member("FindFirst",{VbaValue{}}); });
    expect(450,[&] { clone->member("FindFirst",{}); });
    expect(13,[&] { clone->assign("Bookmark",VbaValue::string("not a bookmark")); });
    expect(94,[&] { clone->assign("Bookmark",VbaValue{}); });
    expect(3159,[&] { clone->assign("Bookmark",VbaValue::byteArray({})); });
    auto damaged=second; damaged.bytes[0]^=1;
    expect(3159,[&] { clone->assign("Bookmark",damaged); });
    auto independent=database->member("OpenRecordset",{VbaValue::string("cursor_values")}).object;
    expect(3159,[&] { independent->assign("Bookmark",second); });
    auto sibling=source->member("Clone",{}).object;
    find(sibling,"FindFirst","id=3"); const auto third=sibling->member("Bookmark",{});
    source->member("Delete",{});
    expect(3021,[&] { (void)clone->item("id"); });
    assert(sibling->item("id").asNumber()==3);
    expect(3159,[&] { source->assign("Bookmark",second); });
    source->assign("Bookmark",third); sibling->member("Edit",{});
    expect(3219,[&] { sibling->member("Clone",{}); });
    expect(3219,[&] { find(sibling,"FindFirst","id=1"); });
    expect(3219,[&] { sibling->assign("Bookmark",third); });
    sibling->assignItem("label",VbaValue::string("updated")); sibling->member("Update",{});
    assert(source->item("id").asNumber()==3 && source->item("label").asText()=="updated");
    source->member("Close",{}); assert(sibling->item("id").asNumber()==3);
    sibling->member("Close",{}); find(clone,"FindFirst","id=1"); assert(clone->item("id").asNumber()==1);
    clone->member("Requery",{}); expect(3159,[&] { clone->assign("Bookmark",third); });
    find(clone,"FindFirst","id=4"); const auto removed=clone->member("Bookmark",{}); clone->member("Delete",{});
    clone->member("AddNew",{}); clone->assignItem("id",VbaValue::numeric(4)); clone->assignItem("label",VbaValue::string("replacement"));
    clone->member("Update",{}); assert(clone->item("id").asNumber()==4);
    expect(3159,[&] { clone->assign("Bookmark",removed); });
    auto ordered=database->member("OpenRecordset",{VbaValue::string("SELECT id,label FROM cursor_values ORDER BY id DESC"),VbaValue::numeric(4)}).object;
    auto orderedClone=ordered->member("Clone",{}).object;
    find(orderedClone,"FindFirst","id=3"); assert(ordered->item("id").asNumber()==4 && orderedClone->item("id").asNumber()==3);
    ordered->assign("Bookmark",orderedClone->member("Bookmark",{})); assert(ordered->item("id").asNumber()==3);
    ordered->member("Close",{}); assert(orderedClone->item("label").asText()=="updated");
    assert(query.exec("SELECT label FROM cursor_values WHERE id=3") && query.next() && query.value(0).toString()=="updated");
}
