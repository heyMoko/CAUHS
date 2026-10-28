@echo off
cd /d "%~dp0"
if not exist node_modules (
  echo Installing project dependencies...
  call npm install
  if errorlevel 1 goto failed
)
echo Open http://127.0.0.1:5173 in your browser.
call npm run dev
goto end
:failed
echo Dependency installation failed. Check Node.js and your network connection.
pause
:end
