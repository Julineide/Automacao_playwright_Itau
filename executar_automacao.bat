@echo off
:: Configura o caminho da pasta e do log
set PASTA_PROJETO=C:\Users\uih34193\Documents\Automacoes.Py\Automacao_Itau
set ARQUIVO_LOG=%PASTA_PROJETO%\historico_execucao.txt

:: Navega ate a pasta
cd /d "%PASTA_PROJETO%"

:: Registra o inicio da execucao no log
echo [%date% %time%] - Iniciando automacao... >> "%ARQUIVO_LOG%"

:: Executa o script e redireciona erros para o mesmo log
"C:\Users\uih34193\AppData\Local\Programs\Python\Python313\python.exe" Main.py >> "%ARQUIVO_LOG%" 2>&1

:: Verifica se o Python retornou erro (0 = sucesso)
if %ERRORLEVEL% EQU 0 (
    echo [%date% %time%] - Fim da execucao: SUCESSO. >> "%ARQUIVO_LOG%"
) else (
    echo [%date% %time%] - Fim da execucao: ERRO (Codigo %ERRORLEVEL%). >> "%ARQUIVO_LOG%"
)

echo -------------------------------------------------- >> "%ARQUIVO_LOG%"