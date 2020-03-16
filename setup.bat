@echo off

::
:: Initial enviroment setting for python automation script.
::


call '%USERPROFILE%\Envs'

:: create virtaul enviroment and install dependencies
call mkvirtualenv testing
call workon testing
call pip install --upgrade pywinauto
call 
call lsvirtualenv
pause