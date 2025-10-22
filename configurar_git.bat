@echo off
title Configuração do Projeto CtrlTransacoesCartao no GitHub
color 0a

echo ==========================================
echo  🚀 CONFIGURANDO GIT PARA O PROJETO CtrlTransacoesCartao
echo ==========================================
echo.

REM === Verifica se o Git está instalado ===
where git >nul 2>nul
if %errorlevel% neq 0 (
    echo ❌ Git não encontrado no sistema.
    echo 👉 Baixe e instale o Git for Windows em:
    echo     https://git-scm.com/download/win
    echo.
    echo Após instalar, execute novamente este arquivo.
    pause
    exit /b
)

REM === Configura nome e e-mail do Git ===
echo.
set /p GITNAME=Digite seu nome completo para o Git: 
set /p GITEMAIL=Digite seu e-mail do GitHub: 
git config --global user.name "%GITNAME%"
git config --global user.email "%GITEMAIL%"
echo ✅ Nome e e-mail configurados no Git.
echo.

REM === Inicializa o repositório ===
cd /d "C:\CtrlTransacoesCartao"
git init
git add .
git commit -m "Versão inicial do projeto CtrlTransacoesCartao"
echo ✅ Repositório Git inicializado e commit criado.
echo.

REM === Mostra instruções para criar o repositório no GitHub ===
echo ==========================================
echo  🧭 PROXIMOS PASSOS:
echo ==========================================
echo 1️⃣ Acesse: https://github.com/new
echo 2️⃣ Crie um repositório chamado: CtrlTransacoesCartao
echo 3️⃣ NÃO adicione README, .gitignore ou License (deixe vazio)
echo 4️⃣ Depois execute os comandos abaixo:
echo.
echo    git remote add origin https://github.com/SEU_USUARIO/CtrlTransacoesCartao.git
echo    git branch -M main
echo    git push -u origin main
echo.
echo ==========================================
echo  💡 Dica: Copie os comandos acima e cole no prompt aqui mesmo.
echo ==========================================

pause
