@echo off
setlocal

rem バッチファイルのある場所(tools)の1つ上をカレントディレクトリにする
cd /d "%~dp0\.."

rem Visual Studioの環境を自動検索
for /f "usebackq tokens=*" %%i in (`"%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -find Common7\Tools\VsDevCmd.bat`) do (
  call "%%i" -arch=x64
  goto :found
)
echo vswhere or VS BuildTools not found.
pause
exit /b 1

:found
echo コンパイルを開始します...

rem srcフォルダの中にあるcppを指定してコンパイル
cl /nologo /std:c++17 /EHsc /LD ^
  /DUNICODE /D_UNICODE /utf-8 /wd4828 ^
  src\MyAssetManager.cpp ^
  user32.lib gdi32.lib comctl32.lib shell32.lib ^
  /Fe:MyAssetManager.dll

REM 出力拡張子を .aux2 に合わせる
if exist MyAssetManager.aux2 del /q MyAssetManager.aux2
rename MyAssetManager.dll MyAssetManager.aux2

REM 一時ファイルのお掃除
if exist MyAssetManager.obj del MyAssetManager.obj
if exist MyAssetManager.exp del MyAssetManager.exp
if exist MyAssetManager.lib del MyAssetManager.lib

echo.
if exist MyAssetManager.aux2 (
    echo [Build OK] MyAssetManager.aux2 の生成に成功しました！
) else (
    echo [Error] ビルドに失敗しました。
)

pause
endlocal