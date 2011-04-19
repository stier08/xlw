@SETLOCAL
@SET LOG_ROOT=%CD%\%2-gcc-
@pushd .
@cd %~p1
make BUILD=DEBUG rebuild >> "%LOG_ROOT%x86-Debug.log" 2>&1 
make BUILD=RELEASE rebuild >> "%LOG_ROOT%x86-Release.log" 2>&1 
@popd
@ENDLOCAL