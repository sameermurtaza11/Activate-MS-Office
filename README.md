# Activate-MS-Office
cd /d %ProgramFiles%\Microsoft Office\Office16<br>
for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16%x"<br>
cscript ospp.vbs /setprt:1688<br>
cscript ospp.vbs /unpkey:6F7TH >nul<br>
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH<br>
cscript ospp.vbs /sethst:s8.uk.to<br>
cscript ospp.vbs /act
