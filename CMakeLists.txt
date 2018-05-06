cmake_minimum_required(VERSION 3.9)

project(rst2pptx
    VERSION 0.4.1
    DESCRIPTION "Generate Powerpoint from Restructured Text"
    Languages NONE
    )

configure_file(setup.cfg.in ${CMAKE_CURRENT_SOURCE_DIR}/setup.cfg)
find_package(PythonInterp)

add_custom_command(OUTPUT ${CMAKE_CURRENT_BINARY_DIR}/${PROJECT_NAME}-${PROJECT_VERSION}-py2-none-any.whl
    COMMAND ${PYTHON_EXECUTABLE} -m pip wheel -w ${CMAKE_CURRENT_BINARY_DIR} ${CMAKE_CURRENT_SOURCE_DIR}
    )
add_custom_target(wheel ALL DEPENDS ${CMAKE_CURRENT_BINARY_DIR}/${PROJECT_NAME}-${PROJECT_VERSION}-py2-none-any.whl)
    
