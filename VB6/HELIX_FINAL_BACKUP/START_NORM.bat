cd C:\HELIX\VirtualDub-1.6.17
call c:\HELIX\capture.bat
cd C:\HELIX\OUTPUT
md C:\HELIX\OUTPUT\%date%
cd C:\HELIX
call 3gp26.bat
cd C:\HELIX
call 3gp82.bat
move C:\HELIX\OUTPUT\26_2g.3gp C:\HELIX\OUTPUT\%date%\26_2g.3gp
move C:\HELIX\OUTPUT\82_3g.3gp C:\HELIX\OUTPUT\%date%\82_3g.3gp
cd  C:\HELIX\OUTPUT\%date%
ftp -i -s:c:\HELIX\ftp.txt 212.152.70.17 > C:\HELIX\LOGS\FTP_LOGS\FTPLOG_%date%.TXT
cd c:\HELIX
