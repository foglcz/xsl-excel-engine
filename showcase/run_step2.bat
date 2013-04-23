@echo off

mkdir step2\output

cd step2\output
"C:\Program Files\Saxonica\SaxonHE9.4N\bin\Transform.exe" -xsl:../blank.xslt -s:../data.xml
cd ..\..

del /Q ".\step2\output\xl\media\blank_image.jpg"
del /Q step2.xlsx

copy step2\blank_image.jpg step2\output\xl\media\blank_image.jpg

cd step2\output
..\..\zip -m -r step2.zip *
move step2.zip ..\..\step2.xlsx
cd ..\..
