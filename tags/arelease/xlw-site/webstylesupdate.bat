
if not %1.==. goto NOTNANDO
set USERNAME=nando
goto DONE

:NOTNANDO
set USERNAME=%1
:DONE


scp -C -r styles/*.css %USERNAME%@shell.sourceforge.net:/home/groups/x/xl/xlw/htdocs/styles

pause
