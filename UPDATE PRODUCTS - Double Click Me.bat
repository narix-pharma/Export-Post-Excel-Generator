@echo off
title Narix — Update Product Database
color 0A
cls

echo.
echo  ==========================================
echo    NARIX — Product Database Updater
echo  ==========================================
echo.

cd /d "%~dp0"

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  [ERROR] Python not found. Start the main app first to install Python.
    pause
    exit /b 1
)

:: Check Excel file exists
if not exist "product_data\PRODUCT_DATA.xlsx" (
    echo  [ERROR] Could not find:
    echo          product_data\PRODUCT_DATA.xlsx
    echo.
    echo  Make sure the Excel file is in the product_data folder.
    pause
    exit /b 1
)

echo  [1/2] Reading product_data\PRODUCT_DATA.xlsx ...
echo.

python update_products.py
if %errorlevel% neq 0 (
    echo.
    echo  [ERROR] Something went wrong. Check the Excel file and try again.
    pause
    exit /b 1
)

echo.
echo  [2/2] Product database updated successfully!
echo.
echo  Restart the app for changes to take effect.
echo  ==========================================
echo.
pause
