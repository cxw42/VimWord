@echo off
:: Install-VimWord.bat: Install VimWord.
:: Copyright (c) 2012--2018 Chris White
:: This file is licensed CC-BY-SA 3.0.
::  2012/11/05  cwhite  Initial version
::  2012/11/13  cwhite  Created from Update-and-respawn
::  2015/07/08  chrisw  Changed to VimWord
::  2015/07/22  chrisw  Added "echo d" just in case
::  2016/10/14  chrisw  Added PATH setting because a user needed it!
::  2017/01/16  chrisw  Added ZIP install check

cls
:: Wipe "Can't run from UNC path" message so users don't get confused
::  Thanks to http://stackoverflow.com/a/9018466/2877364 by
::  http://stackoverflow.com/users/2441/aphoria

path %PATH%;c:\windows\system32
::         ^^^ Because a user had a situation where we needed this!

echo Installer of 2017/01/16

set whereami=%~dp0
:: The path of this bat file, ending with a backslash.
:: We assume the dotm is in the same directory.
::  Thanks to http://stackoverflow.com/a/26564834/2877364 by
::  http://stackoverflow.com/users/2475211/jayro-greybeard

:: TODO update - this fails if whereami includes an ampersand - everything
:: at and after the ampersand is not part of %whereami%

set src="%whereami%VimWord.dotm"
echo Installing from %src%...
if not exist %src% goto :nosource

echo d | xcopy %src% "%appdata%\Microsoft\Word\Startup" /v /f /y
if errorlevel 1 goto error

set src="%whereami%VimWordScratchpad.dotm"
echo Installing from %src%...
if not exist %src% goto :nosource

echo d | xcopy %src% "%appdata%\Microsoft\Word\Startup" /v /f /y
if errorlevel 1 goto error

:: success ::

echo Installed successfully!
goto end

:: failure ::

:error
echo An error occurred
goto end

:: Source file not found ::

:nosource
echo.
echo Could not find %src%.
echo.
echo If you are running this from within the ZIP file, please
echo unzip to a folder, then try again from the folder.
goto end

:end
pause

:: vi: set ts=2 sts=2 sw=2 expandtab ai: ::
