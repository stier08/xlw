
if not %1.==. goto NOTNANDO
set USERNAME=nando
goto DONE

:NOTNANDO
set USERNAME=%1
:DONE


scp -C -r images/*.jpg %USERNAME%@shell.sourceforge.net:/home/groups/x/xl/xlw/htdocs/images

pause
