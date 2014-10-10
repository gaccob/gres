@echo off
..\dep\python.exe ..\conv2pb.py --excel_file=mail.xls --proto=mail.proto --import_path=. --protoc_tool=..\dep\protoc.exe
pause
