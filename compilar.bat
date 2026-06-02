@echo off
taskkill /f /im java.exe >nul 2>&1
del /q *.class >nul 2>&1
".\jdk\bin\javac.exe" -cp "lib/excel/*;." Main.java
if exist Main.class (
    ".\jdk\bin\jar.exe" --create --file=ComparadorFinal.jar --main-class=Main *.class
    echo ==========================================
    echo COMPILADO COM SUCESSO! MESTRE DO NIMBUS!
    echo ==========================================
) else (
    echo ERRO NA COMPILACAO!
)
pause