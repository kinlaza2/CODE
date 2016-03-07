move C:\HELIX\dos_log.txt C:\HELIX\LOGS\DOS_LOGS\DOSLOG_%date%.TXT
cd C:\HELIX\VirtualDub-1.6.17
echo ~~~~~~~~~~~~~~~~ %date% ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ >> C:\HELIX\LOGS\DOS_LOGS\DOSLOG_%date%.TXT
echo START AVI CAPTURE - %time% >> C:\HELIX\LOGS\DOS_LOGS\DOSLOG_%date%.TXT
call c:\HELIX\capture.bat
echo STOP AVI CAPTURE - %time% >> C:\HELIX\LOGS\DOS_LOGS\DOSLOG_%date%.TXT
cd C:\HELIX\OUTPUT
echo MAKE DIRECTORY C:\HELIX\OUTPUT\%date% - %TIME% >> C:\HELIX\LOGS\DOS_LOGS\DOSLOG_%date%.TXT 
md C:\HELIX\OUTPUT\%date%
cd C:\HELIX
echo START ENCODING 26 3gp - %time% >> C:\HELIX\LOGS\DOS_LOGS\DOSLOG_%date%.TXT
call 3gp26.bat
echo FINISH ENCODING 26 3gp  - %time% >> C:\HELIX\LOGS\DOS_LOGS\DOSLOG_%date%.TXT
cd C:\HELIX
echo START ENCODING 82 3gp - %time% >> C:\HELIX\LOGS\DOS_LOGS\DOSLOG_%date%.TXT
call 3gp82.bat
echo FINISH ENCODING 82 3gp - %time% >> C:\HELIX\LOGS\DOS_LOGS\DOSLOG_%date%.TXT
echo MOVE 26 3gp - %time% >> C:\HELIX\LOGS\DOS_LOGS\DOSLOG_%date%.TXT
move C:\HELIX\OUTPUT\26_2g.3gp C:\HELIX\OUTPUT\%date%
echo MOVE 82 3gp - %time% >> C:\HELIX\LOGS\DOS_LOGS\DOSLOG_%date%.TXT
move C:\HELIX\OUTPUT\82_3g.3gp C:\HELIX\OUTPUT\%date%
echo GOTO FOLDER C:\HELIX\OUTPUT\%date% - %TIME% >> C:\HELIX\LOGS\DOS_LOGS\DOSLOG_%date%.TXT
cd  C:\HELIX\OUTPUT\%date%
echo FTP START - %time% >> C:\HELIX\LOGS\DOS_LOGS\DOSLOG_%date%.TXT
ftp -i -s:c:\HELIX\ftp.txt 212.152.70.17 > C:\HELIX\LOGS\FTP_LOGS\FTPLOG_%date%.TXT
echo FTP FINISHED - %time% >> C:\HELIX\LOGS\DOS_LOGS\DOSLOG_%date%.TXT
cd c:\HELIX
echo PROCESS FINISHED - %time% >> C:\HELIX\LOGS\DOS_LOGS\DOSLOG_%date%.TXT
