@echo off
cd /d "%~dp0"
echo Starting Nancy Finds Weekly dashboard...
start "" http://localhost:4321/
node server.js
