rem curl --request POST http://localhost:3000/forms/libreoffice/convert --form files=@test1.xlsx --form files=@test1.docx --form merge=false -o my.zip
rem pause

curl --request POST http://localhost:3000/forms/libreoffice/convert --form files=@test1.xlsx --form files=@test1.docx --form merge=true -o my.pdf
pause
