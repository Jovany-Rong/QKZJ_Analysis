d:
cd D:\dev\python3\全库质检报告检查

del *.spec
rd /s /q build 
rd /s /q dist

pyinstaller -F -c -i "DCTK_Check.ico" "DCTK_Check.py"

xcopy config.ini dist\

del *.spec
rd /s /q build

pause