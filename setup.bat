@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

:: Ensure window stays open even on unexpected errors
if "%~1"=="" (
    cmd /k "%~f0" RUN
    exit /b
)
shift

echo.
echo  AutoCAD MCP Server v5.0 - Setup
echo  ================================
echo.

:: Step 0: Check prerequisites
where git >nul 2>&1
if !errorlevel! neq 0 (
    echo  [INFO] Git hittades inte.
    echo         Git behovs for Claude Code CLI.
    echo.
    echo  Installerar Git via winget...
    winget install Git.Git --accept-source-agreements --accept-package-agreements >nul 2>&1
    if !errorlevel! neq 0 (
        echo  [VARNING] Kunde inte installera Git automatiskt.
        echo            Ladda ner manuellt: https://git-scm.com/downloads/win
        echo.
    ) else (
        echo  [OK] Git installerat. Du kan behova starta om terminalen.
        set "PATH=%ProgramFiles%\Git\cmd;%ProgramFiles%\Git\bin;%PATH%"
    )
) else (
    for /f "delims=" %%i in ('git --version') do echo  [OK] %%i
)

where node >nul 2>&1
if !errorlevel! neq 0 (
    echo  [INFO] Node.js hittades inte.
    echo         Node.js behovs for Claude Code CLI.
    echo.
    echo  Installerar Node.js via winget...
    winget install OpenJS.NodeJS.LTS --accept-source-agreements --accept-package-agreements >nul 2>&1
    if !errorlevel! neq 0 (
        echo  [VARNING] Kunde inte installera Node.js automatiskt.
        echo            Ladda ner manuellt: https://nodejs.org/
        echo.
    ) else (
        echo  [OK] Node.js installerat. Du kan behova starta om terminalen.
        set "PATH=%ProgramFiles%\nodejs;%PATH%"
    )
) else (
    for /f "delims=" %%i in ('node --version') do echo  [OK] Node.js: %%i
)

where claude >nul 2>&1
if !errorlevel! neq 0 (
    echo  [INFO] Claude Code CLI hittades inte.
    where npm >nul 2>&1
    if !errorlevel! equ 0 (
        echo         Installerar Claude Code CLI...
        npm install -g @anthropic-ai/claude-code >nul 2>&1
        if !errorlevel! equ 0 (
            echo  [OK] Claude Code CLI installerat
        ) else (
            echo  [VARNING] Kunde inte installera Claude Code CLI.
            echo            Kor manuellt: npm install -g @anthropic-ai/claude-code
        )
    ) else (
        echo         Installera Node.js forst, kor sedan: npm install -g @anthropic-ai/claude-code
    )
) else (
    for /f "delims=" %%i in ('claude --version 2^>nul') do echo  [OK] Claude Code: %%i
)

where codex >nul 2>&1
if !errorlevel! neq 0 (
    echo  [INFO] Codex CLI hittades inte.
    where npm >nul 2>&1
    if !errorlevel! equ 0 (
        echo         Installerar Codex CLI...
        npm install -g @openai/codex >nul 2>&1
        if !errorlevel! equ 0 (
            echo  [OK] Codex CLI installerat
        ) else (
            echo  [VARNING] Kunde inte installera Codex CLI.
            echo            Kor manuellt: npm install -g @openai/codex
        )
    ) else (
        echo         Installera Node.js forst, kor sedan: npm install -g @openai/codex
    )
) else (
    for /f "delims=" %%i in ('codex --version 2^>nul') do echo  [OK] Codex CLI: %%i
)
echo.

:: Step 1: Find our own directory (where this .bat lives = project root)
set "PROJECT_DIR=%~dp0"
:: Remove trailing backslash
if "%PROJECT_DIR:~-1%"=="\" set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"

if not exist "%PROJECT_DIR%\pyproject.toml" (
    echo  [FEL] Kan inte hitta pyproject.toml i %PROJECT_DIR%
    goto :error
)
if not exist "%PROJECT_DIR%\src\autocad_mcp\__init__.py" (
    echo  [FEL] Kan inte hitta src/autocad_mcp/ i %PROJECT_DIR%
    goto :error
)
echo  [OK] Projekt: %PROJECT_DIR%

:: Step 2: Check/install uv
where uv >nul 2>&1
if !errorlevel! neq 0 (
    set "PATH=%USERPROFILE%\.local\bin;%USERPROFILE%\.cargo\bin;%PATH%"
    where uv >nul 2>&1
)
if !errorlevel! neq 0 (
    echo.
    echo  uv hittades inte. Installerar...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "irm https://astral.sh/uv/install.ps1 | iex"
    if !errorlevel! neq 0 (
        echo  [FEL] Kunde inte installera uv.
        echo  Installera manuellt: https://docs.astral.sh/uv/getting-started/installation/
        goto :error
    )
    set "PATH=%USERPROFILE%\.local\bin;%USERPROFILE%\.cargo\bin;%PATH%"
    where uv >nul 2>&1
    if !errorlevel! neq 0 (
        echo  [FEL] uv installerades men hittades inte i PATH.
        echo  Stang och oppna terminalen, kor sedan setup.bat igen.
        goto :error
    )
)
for /f "delims=" %%i in ('where uv') do set "UV_PATH=%%i"
echo  [OK] uv: %UV_PATH%

:: Step 3: Pre-cache dependencies
echo.
echo  Forbereder Python-beroenden...

:: Use local .venv to avoid SharePoint sync conflicts between colleagues
set "LOCAL_VENV=%LOCALAPPDATA%\autocad-mcp\.venv"
set "UV_PROJECT_ENVIRONMENT=%LOCAL_VENV%"

:: Clean up stale .venv from SharePoint folder if it exists (legacy cleanup)
if exist "%PROJECT_DIR%\.venv" (
    echo  Tar bort gammal .venv fran delad mapp...
    rmdir /s /q "%PROJECT_DIR%\.venv" >nul 2>&1
)

:: Pre-cache by creating the local .venv with all dependencies
uv --directory "%PROJECT_DIR%" run python -c "print('ok')" >nul 2>&1
if !errorlevel! neq 0 (
    echo  [VARNING] uv pre-cache misslyckades. MCP-servern laddar beroenden vid forsta start.
) else (
    echo  [OK] Beroenden cachade (mcp, ezdxf, matplotlib, Pillow, structlog, openpyxl, pywin32)
)

:: Step 4: Configure Claude Code CLI + Claude Desktop
echo.
echo  Konfigurerar MCP-server...

:: Pass paths via environment variables (avoids quoting issues with spaces)
set "ACAD_UV=%UV_PATH%"
set "ACAD_DIR=%PROJECT_DIR%"

:: Write PowerShell script to temp
set "PS_SCRIPT=%TEMP%\autocad_configure_mcp.ps1"
> "%PS_SCRIPT%" echo $uv = $env:ACAD_UV
>> "%PS_SCRIPT%" echo $dir = $env:ACAD_DIR
>> "%PS_SCRIPT%" echo $ea = @('--directory',$dir,'run','python','-m','autocad_mcp')
>> "%PS_SCRIPT%" echo $localVenv = Join-Path $env:LOCALAPPDATA 'autocad-mcp\.venv'
>> "%PS_SCRIPT%" echo $envVars = [pscustomobject]@{ AUTOCAD_MCP_BACKEND = 'auto'; UV_PROJECT_ENVIRONMENT = $localVenv }
>> "%PS_SCRIPT%" echo $entry = [pscustomobject]@{ command = $uv; args = $ea; env = $envVars }
>> "%PS_SCRIPT%" echo function Fix-Json($t) { return [regex]::Replace($t, ',(\s*[\]\}])', '$1') }
>> "%PS_SCRIPT%" echo function Update-Config($path, $key, $label) {
>> "%PS_SCRIPT%" echo   try {
>> "%PS_SCRIPT%" echo     $cfg = $null
>> "%PS_SCRIPT%" echo     if (Test-Path $path) {
>> "%PS_SCRIPT%" echo       Write-Host "  [$label] $path"
>> "%PS_SCRIPT%" echo       try {
>> "%PS_SCRIPT%" echo         $raw = Get-Content $path -Raw -Encoding UTF8
>> "%PS_SCRIPT%" echo         $cfg = Fix-Json $raw ^| ConvertFrom-Json
>> "%PS_SCRIPT%" echo       } catch {
>> "%PS_SCRIPT%" echo         Copy-Item $path "$path.bak" -Force
>> "%PS_SCRIPT%" echo         Write-Host "  [$label] JSON trasig, backup sparad" -ForegroundColor Yellow
>> "%PS_SCRIPT%" echo       }
>> "%PS_SCRIPT%" echo     }
>> "%PS_SCRIPT%" echo     if (-not $cfg) {
>> "%PS_SCRIPT%" echo       $dir = Split-Path $path -Parent
>> "%PS_SCRIPT%" echo       if (-not (Test-Path $dir)) { New-Item -ItemType Directory $dir ^| Out-Null }
>> "%PS_SCRIPT%" echo       $cfg = [pscustomobject]@{}
>> "%PS_SCRIPT%" echo     }
>> "%PS_SCRIPT%" echo     if (-not $cfg.mcpServers) {
>> "%PS_SCRIPT%" echo       $cfg ^| Add-Member -NotePropertyName mcpServers -NotePropertyValue ([pscustomobject]@{}) -Force
>> "%PS_SCRIPT%" echo     }
>> "%PS_SCRIPT%" echo     $cfg.mcpServers ^| Add-Member -NotePropertyName $key -NotePropertyValue $entry -Force
>> "%PS_SCRIPT%" echo     $json = Fix-Json ($cfg ^| ConvertTo-Json -Depth 10)
>> "%PS_SCRIPT%" echo     [IO.File]::WriteAllText($path, $json, [Text.UTF8Encoding]::new($false))
>> "%PS_SCRIPT%" echo     Write-Host "  [OK] $label : $path"
>> "%PS_SCRIPT%" echo   } catch {
>> "%PS_SCRIPT%" echo     Write-Host "  [FEL] $label : $_" -ForegroundColor Red
>> "%PS_SCRIPT%" echo   }
>> "%PS_SCRIPT%" echo }
>> "%PS_SCRIPT%" echo Update-Config (Join-Path $env:USERPROFILE '.claude.json') 'autocad-mcp' 'Claude Code CLI'
>> "%PS_SCRIPT%" echo $found = $false
>> "%PS_SCRIPT%" echo $std = Join-Path $env:APPDATA 'Claude\claude_desktop_config.json'
>> "%PS_SCRIPT%" echo if (Test-Path $std) { Update-Config $std 'AutoCAD MCP Server' 'Claude Desktop'; $found = $true }
>> "%PS_SCRIPT%" echo $dirs = Get-ChildItem "$env:LOCALAPPDATA\Packages\Claude_*" -Directory -ErrorAction SilentlyContinue
>> "%PS_SCRIPT%" echo foreach ($d in $dirs) {
>> "%PS_SCRIPT%" echo   $sp = Join-Path $d.FullName 'LocalCache\Roaming\Claude\claude_desktop_config.json'
>> "%PS_SCRIPT%" echo   if (Test-Path $sp) { Update-Config $sp 'AutoCAD MCP Server' 'Claude Desktop (Store)'; $found = $true }
>> "%PS_SCRIPT%" echo }
>> "%PS_SCRIPT%" echo if (-not $found) { Update-Config $std 'AutoCAD MCP Server' 'Claude Desktop' }

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"

if !errorlevel! neq 0 (
    echo.
    echo  [VARNING] Kunde inte uppdatera config-filer automatiskt.
    echo  Kontrollera felmeddelanden ovan.
)

del "%PS_SCRIPT%" >nul 2>&1

:: Step 5: Configure OpenAI Codex CLI (if installed)
where codex >nul 2>&1
if !errorlevel! equ 0 (
    call :configure_codex
) else (
    echo.
    echo  [INFO] Codex CLI hittades inte, hoppar over.
)

:: Step 6: Auto-load mcp_dispatch.lsp in AutoCAD
echo.
echo  Konfigurerar AutoCAD LISP auto-load...

powershell -NoProfile -ExecutionPolicy Bypass -File "%PROJECT_DIR%\config\autoload_lisp.ps1"

if !errorlevel! neq 0 (
    echo.
    echo  [VARNING] Kunde inte konfigurera AutoCAD LISP auto-load.
    echo  Du kan ladda mcp_dispatch.lsp manuellt via APPLOAD i AutoCAD.
)

:: Step 7: Test AutoCAD connection
echo.
echo  Testar AutoCAD MCP-server...
uv --directory "%PROJECT_DIR%" run python -c "from autocad_mcp.config import detect_backend; b = detect_backend(); print('[OK] Backend: ' + b)" 2>nul
if !errorlevel! neq 0 (
    echo  [INFO] Kunde inte testa servern. Kontrollera att beroenden ar installerade.
    echo         Kor: uv --directory "%PROJECT_DIR%" run python -m autocad_mcp
)

:: Done
echo.
echo  ================================
echo  Setup klar!
echo.
echo  uv:       %UV_PATH%
echo  projekt:  %PROJECT_DIR%
echo.
echo  Backend-val (miljovariabler):
echo    AUTOCAD_MCP_BACKEND = auto ^| file_ipc ^| com ^| ezdxf
echo    AUTOCAD_MCP_CAD_TYPE = autocad ^| zwcad ^| gcad ^| bricscad
echo.
echo  File IPC (AutoCAD LT): mcp_dispatch.lsp laddas automatiskt vid start
echo  COM: starta AutoCAD/ZWCAD/GstarCAD/BricsCAD fore anslutning
echo  ezdxf: ingen CAD-instans behovs (headless DXF)
echo.
echo  Starta om AutoCAD, Claude Code, Claude Desktop och/eller Codex.
echo  Verifiera med: "Check AutoCAD status" i Claude Code.
echo.
echo  Avinstallera LISP auto-load:
echo    powershell -File "%PROJECT_DIR%\config\autoload_lisp.ps1" -Uninstall
echo.
pause
exit /b 0

:configure_codex
echo.
echo  Konfigurerar Codex CLI...
codex mcp remove autocad_mcp >nul 2>&1
codex mcp add autocad_mcp -- "%UV_PATH%" --directory "%PROJECT_DIR%" run python -m autocad_mcp >nul 2>&1
if !errorlevel! equ 0 (
    echo  [OK] Codex CLI
) else (
    echo  [VARNING] Kunde inte konfigurera Codex CLI
)
exit /b

:error
echo.
echo  Setup avbrots. Se felmeddelanden ovan.
echo.
pause
exit /b 1
