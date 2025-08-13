@echo off
setlocal

echo.
echo =================================================================
echo           INICIANDO PROCESSO DE BUILD ROBUSTO
echo =================================================================
echo.

echo [FASE 1/4] Limpando builds antigos e ambiente virtual...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "venv" rmdir /s /q venv
if exist "Codificacao_Estacao.spec" del "Codificacao_Estacao.spec"
echo Limpeza concluida.
echo.

echo [FASE 2/4] Criando novo ambiente virtual e instalando dependencias...
python -m venv venv
if %errorlevel% neq 0 (
    echo ERRO: Nao foi possivel criar o ambiente virtual. Verifique sua instalacao do Python.
    pause
    exit /b
)
call venv\Scripts\activate

python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller pyinstaller-hooks-contrib
echo Dependencias instaladas.
echo.

echo [FASE 3/4] Gerando executavel com PyInstaller (isso pode demorar)...
pyinstaller --noconfirm ^
    --onefile ^
    --windowed ^
    --name "Codificacao_Estacao" ^
    --icon="assets/pluviometer.ico" ^
    --add-data="assets;assets" ^
    --add-data="mdb/template.mdb;mdb" ^
    --collect-all numpy ^
    --copy-metadata pandas ^
    --hidden-import="pywin32" ^
    --log-level=INFO ^
    Codificacao_Estacao_GUI.py

if %errorlevel% neq 0 (
    echo.
    echo !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    echo    ERRO: PyInstaller falhou ao gerar o executavel.
    echo    Verifique o log de erros acima.
    echo !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    pause
    exit /b
)

echo.
echo [FASE 4/4] Finalizando...
echo.
echo =================================================================
echo      BUILD CONCLUIDO COM SUCESSO!
echo =================================================================
echo O executavel "Codificacao_Estacao.exe" esta na pasta 'dist'.
echo.
pause
endlocal