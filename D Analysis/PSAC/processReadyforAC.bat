@ECHO OFF
SET ThisScriptsDirectory=C:\Users\DEPPOPP1\OneDrive - EY\Dokumente\GitHub\AC\Adressabgleich\D Analysis\PSAC\PSAP
SET PowerShellScriptPath=C:\Users\DEPPOPP1\OneDrive - EY\Dokumente\GitHub\AC\Adressabgleich\D Analysis\PSAC\processReadyforAC.ps1
%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell -NoProfile -ExecutionPolicy Bypass -Command "& '%PowerShellScriptPath%'";
PAUSE