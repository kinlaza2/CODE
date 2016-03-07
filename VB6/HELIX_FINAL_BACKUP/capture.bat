echo CAPTURE AVI START %date% %time%
cd C:\HELIX\VirtualDub-1.6.17
vdub /capture  /capfile D:\HELIX\%date%.avi /capstart 120
echo CAPTURE AVI END SUCCESFULLY %date% %time%
cd C:\HELIX