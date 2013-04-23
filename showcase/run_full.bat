@echo off

mkdir fullExample\output

cd fullExample\output
"C:\Program Files\Saxonica\SaxonHE9.4N\bin\Transform.exe" -xsl:../template.xslt -s:../data.xml
cd ..\..

del /Q ".\fullExample\output\xl\media\foglcz.png"
del /Q ".\fullExample\output\xl\printerSettings\Example_printer.bin"
del /Q ".\fullExample\output\xl\Example_vba.bin"
del /Q fullExample.xlsm

copy fullExample\foglcz.png fullExample\output\xl\media\foglcz.png
copy fullExample\Example_printer.bin fullExample\output\xl\printerSettings\Example_printer.bin
copy fullExample\Example_vba.bin fullExample\output\xl\Example_vba.bin

cd fullExample\output
..\..\zip -r -m fullExample.zip *
move fullExample.zip ..\..\fullExample.xlsm
cd ..\..
