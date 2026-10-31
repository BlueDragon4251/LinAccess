cmake_minimum_required(VERSION 3.21)
project(LinAccess VERSION 0.7.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
add_library(linaccess_runtime STATIC src/vba_runtime.cpp src/vba_expression.cpp src/vba_currency.cpp src/vba_analyzer.cpp src/csv.cpp)
target_include_directories(linaccess_runtime PUBLIC src)

option(LINACCESS_BUILD_GUI "Build the native Qt desktop application" ON)
if(LINACCESS_BUILD_GUI)
    set(CMAKE_AUTOMOC ON)
    find_package(Qt6 6.4 REQUIRED COMPONENTS Widgets Sql PrintSupport)
    add_library(linaccess_data STATIC
        src/access_catalog.cpp src/access_table_import.cpp src/access_query.cpp src/access_format.cpp
        src/dao_database.cpp src/sqlite_execution.cpp src/jet_sql_functions.cpp src/vba_source_codec.cpp
        src/vba_qt_value.cpp src/vba_module_context.cpp src/access_form_import.cpp)
    target_link_libraries(linaccess_data PUBLIC linaccess_runtime Qt6::Core Qt6::Sql)
    target_include_directories(linaccess_data PUBLIC src)
    add_executable(linaccess

        src/main.cpp
        src/mainwindow.cpp
        src/mainwindow.hpp
        src/access_inspector.cpp
        src/access_inspector.hpp

        src/access_catalog.hpp

        src/access_table_import.hpp

        src/access_form_import.hpp
        src/access_form_preview.cpp src/native_send_keys.cpp src/access_form_objects.cpp src/access_layout_print.cpp
        src/access_form_preview.hpp
        src/bridge_importer.cpp
        src/bridge_importer.hpp

        src/vba_analyzer.hpp

        src/vba_runtime.hpp

        src/vba_expression.hpp

        src/vba_source_codec.hpp

        src/csv.hpp
    )
    target_link_libraries(linaccess PRIVATE linaccess_runtime linaccess_data Qt6::Widgets Qt6::Sql Qt6::PrintSupport)
    # Cloud workspaces can use umask 0077. Installed system directories must
    # remain traversable by ordinary desktop users in the generated package.
    set(CMAKE_INSTALL_DEFAULT_DIRECTORY_PERMISSIONS
        OWNER_READ OWNER_WRITE OWNER_EXECUTE GROUP_READ GROUP_EXECUTE WORLD_READ WORLD_EXECUTE)
    set(CPACK_INSTALL_DEFAULT_DIRECTORY_PERMISSIONS ${CMAKE_INSTALL_DEFAULT_DIRECTORY_PERMISSIONS})
    foreach(directory IN ITEMS "." "bin" "share" "share/applications")
        install(DIRECTORY DESTINATION "${directory}" DIRECTORY_PERMISSIONS
            OWNER_READ OWNER_WRITE OWNER_EXECUTE GROUP_READ GROUP_EXECUTE WORLD_READ WORLD_EXECUTE)
    endforeach()
    install(TARGETS linaccess RUNTIME DESTINATION bin)
    install(FILES packaging/linaccess.desktop DESTINATION share/applications)

    set(CPACK_PACKAGE_NAME "linaccess")
    set(CPACK_PACKAGE_VENDOR "LinAccess Community")
    set(CPACK_PACKAGE_CONTACT "LinAccess Community")
    set(CPACK_PACKAGE_DESCRIPTION_SUMMARY "Native Linux database designer (early prototype)")
    set(CPACK_DEBIAN_PACKAGE_SHLIBDEPS ON)
    set(CPACK_DEBIAN_PACKAGE_DEPENDS "libqt6sql6-sqlite, mdbtools (>= 1.0)")
    set(CPACK_GENERATOR "DEB")
    include(CPack)
endif()

include(CTest)
if(BUILD_TESTING)
    add_executable(linaccess_core_tests tests/core_tests.cpp    )
    target_include_directories(linaccess_core_tests PRIVATE src)
    target_link_libraries(linaccess_core_tests PRIVATE linaccess_runtime)
    add_test(NAME core_tests COMMAND linaccess_core_tests)
    if(LINACCESS_BUILD_GUI)
        find_package(Qt6 6.4 REQUIRED COMPONENTS Test)
        add_executable(linaccess_dao_tests tests/dao_tests.cpp
             )
        target_include_directories(linaccess_dao_tests PRIVATE src)
        target_link_libraries(linaccess_dao_tests PRIVATE linaccess_runtime linaccess_data Qt6::Core Qt6::Sql)
        add_test(NAME dao_execution COMMAND linaccess_dao_tests)
        add_executable(linaccess_query_model_tests tests/query_model_tests.cpp)
        target_link_libraries(linaccess_query_model_tests PRIVATE linaccess_data)
        add_test(NAME access_query_model COMMAND linaccess_query_model_tests)
        add_executable(linaccess_format_tests tests/access_format_tests.cpp)
        target_link_libraries(linaccess_format_tests PRIVATE linaccess_data)
        add_test(NAME access_date_formats COMMAND linaccess_format_tests)
        add_executable(linaccess_concat_tests tests/concat_tests.cpp)
        target_link_libraries(linaccess_concat_tests PRIVATE linaccess_data)
        add_test(NAME jet_concatenation COMMAND linaccess_concat_tests)
        add_executable(linaccess_sql_object_tests tests/sql_object_tests.cpp)
        target_link_libraries(linaccess_sql_object_tests PRIVATE linaccess_data)
        add_test(NAME sql_object_values COMMAND linaccess_sql_object_tests)
        add_executable(linaccess_projection_tests tests/projection_tests.cpp)
        target_link_libraries(linaccess_projection_tests PRIVATE linaccess_data)
        add_test(NAME jet_projection_names COMMAND linaccess_projection_tests)
        set_tests_properties(access_date_formats PROPERTIES ENVIRONMENT "TZ=Europe/Berlin")
        add_executable(linaccess_table_import_tests tests/access_table_import_tests.cpp
              )
        target_include_directories(linaccess_table_import_tests PRIVATE src)
        target_link_libraries(linaccess_table_import_tests PRIVATE linaccess_runtime linaccess_data Qt6::Core Qt6::Sql)
        add_test(NAME table_import_contracts COMMAND linaccess_table_import_tests)
        add_executable(linaccess_mdb_tests
            tests/mdb_tests.cpp

            src/access_catalog.hpp

            src/access_table_import.hpp

            src/csv.hpp)
        target_include_directories(linaccess_mdb_tests PRIVATE src)
        target_link_libraries(linaccess_mdb_tests PRIVATE linaccess_runtime linaccess_data Qt6::Core Qt6::Sql)
        set(LINACCESS_REAL_MDB_FILE "" CACHE FILEPATH "Optional real MDB regression test path (not distributed)")
        if(LINACCESS_REAL_MDB_FILE AND EXISTS "${LINACCESS_REAL_MDB_FILE}")
            add_test(NAME mdb_real_fixture COMMAND linaccess_mdb_tests "${LINACCESS_REAL_MDB_FILE}")
            add_executable(linaccess_mdb_inspector_gui_tests
                tests/mdb_inspector_gui_tests.cpp
                src/access_inspector.cpp src/access_inspector.hpp
                 src/access_catalog.hpp
                 src/access_table_import.hpp
                 src/csv.hpp)
            target_include_directories(linaccess_mdb_inspector_gui_tests PRIVATE src)
            target_link_libraries(linaccess_mdb_inspector_gui_tests PRIVATE linaccess_runtime linaccess_data Qt6::Widgets Qt6::Sql)
            add_test(NAME mdb_inspector_gui_tests
                COMMAND linaccess_mdb_inspector_gui_tests "${LINACCESS_REAL_MDB_FILE}")
            set_tests_properties(mdb_inspector_gui_tests PROPERTIES ENVIRONMENT "QT_QPA_PLATFORM=offscreen")
        endif()
        add_executable(linaccess_bridge_gui_tests
            tests/bridge_gui_tests.cpp
            src/bridge_importer.cpp
            src/bridge_importer.hpp

            src/vba_source_codec.hpp

            src/vba_runtime.hpp

            src/vba_expression.hpp

            src/vba_analyzer.hpp
        )
        target_include_directories(linaccess_bridge_gui_tests PRIVATE src)
        target_link_libraries(linaccess_bridge_gui_tests PRIVATE linaccess_runtime linaccess_data Qt6::Widgets Qt6::Sql)
        target_compile_definitions(linaccess_bridge_gui_tests PRIVATE
            LINACCESS_FIXTURE_DIRECTORY="${CMAKE_CURRENT_SOURCE_DIR}/examples/bridge_sample")
        add_test(NAME bridge_gui_tests COMMAND linaccess_bridge_gui_tests)
        set_tests_properties(bridge_gui_tests PROPERTIES ENVIRONMENT "QT_QPA_PLATFORM=offscreen")
        set(LINACCESS_REAL_BRIDGE_DIR "" CACHE PATH "Optional trusted real Access Bridge export directory")
        if(LINACCESS_REAL_BRIDGE_DIR AND EXISTS "${LINACCESS_REAL_BRIDGE_DIR}/manifest.json")
            add_executable(linaccess_real_bridge_tests
                tests/real_bridge_tests.cpp
                src/bridge_importer.cpp  src/bridge_importer.hpp
                 src/vba_source_codec.hpp
                 src/vba_analyzer.hpp
                 src/vba_runtime.hpp
                 src/vba_expression.hpp)
            target_sources(linaccess_real_bridge_tests PRIVATE
                 src/access_form_import.hpp)
            target_include_directories(linaccess_real_bridge_tests PRIVATE src)
            target_link_libraries(linaccess_real_bridge_tests PRIVATE linaccess_runtime linaccess_data Qt6::Widgets Qt6::Sql)
            add_test(NAME real_access_bridge
                COMMAND linaccess_real_bridge_tests "${LINACCESS_REAL_BRIDGE_DIR}" 26 4 1 19 2 0)
            set_tests_properties(real_access_bridge PROPERTIES ENVIRONMENT "QT_QPA_PLATFORM=offscreen")
        endif()
        set(LINACCESS_COMPLEX_BRIDGE_DIR "" CACHE PATH "Optional trusted complex real Access Bridge export directory")
        if(LINACCESS_COMPLEX_BRIDGE_DIR AND EXISTS "${LINACCESS_COMPLEX_BRIDGE_DIR}/manifest.json")
            if(NOT TARGET linaccess_real_bridge_tests)
                add_executable(linaccess_real_bridge_tests
                    tests/real_bridge_tests.cpp
                    src/bridge_importer.cpp  src/bridge_importer.hpp
                     src/vba_source_codec.hpp
                     src/vba_analyzer.hpp
                     src/vba_runtime.hpp
                     src/vba_expression.hpp
                     src/access_form_import.hpp)
                target_include_directories(linaccess_real_bridge_tests PRIVATE src)
                target_link_libraries(linaccess_real_bridge_tests PRIVATE linaccess_runtime linaccess_data Qt6::Widgets Qt6::Sql)
            endif()
            add_executable(linaccess_vba_corpus_audit tests/vba_corpus_audit.cpp
                   )
            target_include_directories(linaccess_vba_corpus_audit PRIVATE src)
            target_link_libraries(linaccess_vba_corpus_audit PRIVATE linaccess_runtime linaccess_data Qt6::Core)
            add_test(NAME kassa_vba_inventory COMMAND linaccess_vba_corpus_audit
                "${LINACCESS_COMPLEX_BRIDGE_DIR}" "${CMAKE_CURRENT_BINARY_DIR}/private-vba-inventory.json")
            add_test(NAME complex_access_bridge
                COMMAND linaccess_real_bridge_tests "${LINACCESS_COMPLEX_BRIDGE_DIR}" 192 47 9 100 32 4)
            set_tests_properties(complex_access_bridge PROPERTIES ENVIRONMENT "QT_QPA_PLATFORM=offscreen")
            add_executable(linaccess_access_layout_tests
                tests/access_form_layout_tests.cpp tests/message_box_tests.cpp tests/close_command_tests.cpp tests/document_lifetime_tests.cpp tests/dynamic_combo_tests.cpp tests/active_document_commands_tests.cpp tests/send_keys_tests.cpp tests/calculated_cache_tests.cpp tests/module_state_tests.cpp tests/calendar_tests.cpp tests/filter_tests.cpp tests/form_event_tests.cpp tests/subform_tests.cpp tests/form_object_tests.cpp tests/original_control_tests.cpp tests/report_event_tests.cpp tests/report_corpus_tests.cpp
                 src/access_form_import.hpp

                src/access_form_preview.cpp src/native_send_keys.cpp src/access_form_objects.cpp src/access_layout_print.cpp src/access_form_preview.hpp
                 src/vba_source_codec.hpp
                 src/vba_runtime.hpp
                 src/vba_expression.hpp
                 src/vba_analyzer.hpp)
            target_include_directories(linaccess_access_layout_tests PRIVATE src)
            target_link_libraries(linaccess_access_layout_tests PRIVATE linaccess_runtime linaccess_data Qt6::Widgets Qt6::Sql Qt6::Test Qt6::PrintSupport)
            add_test(NAME access_2003_layouts
                COMMAND linaccess_access_layout_tests "${LINACCESS_COMPLEX_BRIDGE_DIR}")
            set_tests_properties(access_2003_layouts PROPERTIES ENVIRONMENT "QT_QPA_PLATFORM=offscreen")
        endif()
        set(LINACCESS_COMPLEX_MDB_FILE "" CACHE FILEPATH "Optional complex real Access MDB table fixture")
        if(LINACCESS_COMPLEX_MDB_FILE AND EXISTS "${LINACCESS_COMPLEX_MDB_FILE}")
            if(LINACCESS_COMPLEX_BRIDGE_DIR AND EXISTS "${LINACCESS_COMPLEX_BRIDGE_DIR}/manifest.json")
                add_executable(linaccess_query_corpus_tests tests/query_corpus_tests.cpp

                     )
                target_link_libraries(linaccess_query_corpus_tests PRIVATE linaccess_runtime linaccess_data Qt6::Core Qt6::Sql)
                add_test(NAME kassa_query_inventory COMMAND linaccess_query_corpus_tests
                    "${LINACCESS_COMPLEX_BRIDGE_DIR}" "${LINACCESS_COMPLEX_MDB_FILE}"
                    "${CMAKE_CURRENT_BINARY_DIR}/private-query-inventory.json")
            endif()
            if(TARGET linaccess_access_layout_tests)
                add_test(NAME real_kassa_dao_event COMMAND linaccess_access_layout_tests
                    "${LINACCESS_COMPLEX_BRIDGE_DIR}" "${LINACCESS_COMPLEX_MDB_FILE}")
                set_tests_properties(real_kassa_dao_event PROPERTIES ENVIRONMENT "QT_QPA_PLATFORM=offscreen")
            endif()
            add_executable(linaccess_complex_mdb_table_tests
                tests/complex_mdb_table_tests.cpp
                 src/access_catalog.hpp
                 src/access_table_import.hpp
                 src/csv.hpp)
            target_include_directories(linaccess_complex_mdb_table_tests PRIVATE src)
            target_link_libraries(linaccess_complex_mdb_table_tests PRIVATE linaccess_runtime linaccess_data Qt6::Sql Qt6::Core)
            add_test(NAME complex_real_mdb_table_copy
                COMMAND linaccess_complex_mdb_table_tests "${LINACCESS_COMPLEX_MDB_FILE}")
        endif()
    endif()
endif()

