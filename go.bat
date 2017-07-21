@echo off

echo 正在加密，请稍后....
echo path:%~dp0

set base=%~dp0

set class=%base%\bin
set libs=%base%\lib

set class_path=%class%;%libs%\poi-3.16.jar;%libs%\poi-ooxml-3.16.jar;%libs%\poi-ooxml-schemas-3.16.jar;%libs%\poi-scratchpad-3.16.jar;%libs%\xmlbeans-2.6.0.jar;%libs%\dom4j-2.0.1.jar;%libs%\dom4j-2.0.1.jar;

java -classpath %class_path% com.core.cbx.amos.ExcelParser
@pause