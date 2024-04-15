REM  This script adds the members of the MS GG Ext Server Support group to
REM  the local Administrators Group on servers that are in a DMZ.
REM  Joe McConnell is also added
REM  The initial password for these users will "Password1!"
REM  Written by Jim Rigby   Date:  04/30/2009


net user zz148794 Password1! /fullname:"Donald Burt" /add
net user zztorregr Password1! /fullname:"Gregory Torres" /add
net user zzstjohro2 Password1! /fullname:"Robert St.John" /add
net user zzs360754 Password1! /fullname:"Aaron Mais" /add
net user zzs360045 Password1! /fullname:"Darren Briscoe" /add
net user zzrigbyja Password1! /fullname:"Jim Rigby" /add
net user zzphamhe1 Password1! /fullname:"Helen Pham" /add
net user zzPeterDi Password1! /fullname:"Dixie Peterson" /add
net user zzpattech4 Password1! /fullname:"Cheryl Patterson" /add
net user zzfaroona Password1! /fullname:"Nayyar Farooq" /add
net user zzcaohu Password1! /fullname:"Hung Cao" /add
net user zzS819733 Password1! /fullname:"Joseph McConnell" /add

net localgroup administrators zz148794 /add
net localgroup administrators zztorregr /add
net localgroup administrators zzstjohro2 /add
net localgroup administrators zzs360754 /add
net localgroup administrators zzs360045 /add
net localgroup administrators zzrigbyja /add
net localgroup administrators zzphamhe1 /add
net localgroup administrators zzPeterDi /add
net localgroup administrators zzpattech4 /add
net localgroup administrators zzfaroona /add
net localgroup administrators zzcaohu /add
net localgroup administrators zzS819733 /add

net localgroup users
pause
net localgroup administrators
pause