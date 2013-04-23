@echo off

mkdir step1\output

cd step1\output
"C:\Program Files\Saxonica\SaxonHE9.4N\bin\Transform.exe" -xsl:../blank.xslt -s:../data.xml
cd ..\..

del /Q ".\step1\output\xl\media\blank_image.jpg"
del /Q step1.xlsx

copy step1\blank_image.jpg step1\output\xl\media\blank_image.jpg

cd step1\output
..\..\zip -m -r step1.zip *
move step1.zip ..\..\step1.xlsx
cd ..\..
