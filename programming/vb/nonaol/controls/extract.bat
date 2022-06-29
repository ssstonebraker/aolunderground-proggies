@echo off 
setlocal enableextensions enabledelayedexpansion
SETLOCAL

SET OUTDIR="."
SET SRCDIR="."

for %%a in (%SRCDIR%\*.zip) do (
	set file=%%~na
	set odir=!file:DFX-=!
	7z.exe e %SRCDIR%\%%~na.zip -o%OUTDIR%\!odir! *.* -y -r
)

