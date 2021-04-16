@echo off

set /a "x = 0"

:bucle
         if %x% leq 5 (
             
                  set /a "x = x +1"
                  copy Autoduplicar.cmd %x%Autoduplicar.cmd
 
goto :bucle)