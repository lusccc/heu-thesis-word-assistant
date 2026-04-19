@echo off
REM Compile thesis.qmd and apply post-processing styles

set OUTPUT_DIR=output
set OUTPUT_BASENAME=thesis.docx
set OUTPUT_DOCX=%OUTPUT_DIR%\%OUTPUT_BASENAME%

if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

echo Compiling thesis.qmd -^> %OUTPUT_DOCX% ...
quarto render thesis.qmd --to docx --output "%OUTPUT_BASENAME%"
if errorlevel 1 (
    echo Compile failed!
    exit /b 1
)
move /y "%OUTPUT_BASENAME%" "%OUTPUT_DOCX%" >nul

echo Applying post-processing styles...
python scripts\thesis_post_process.py "%OUTPUT_DOCX%" thesis.qmd
if errorlevel 1 (
    echo Post-processing failed!
    exit /b 1
)

echo Opening %OUTPUT_DOCX% in Word ...
start "" "%OUTPUT_DOCX%"

echo Done!
