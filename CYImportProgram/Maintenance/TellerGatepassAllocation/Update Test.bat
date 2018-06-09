set filename=CCRAllocation.dll
set path32=%windir%\System32
set path64=%windir%\SysWOW64
set cdir="%~dp0"

copy %cdir%%filename% %path32% /y
copy %cdir%%filename% %path64% /y

regsvr32 %filename%