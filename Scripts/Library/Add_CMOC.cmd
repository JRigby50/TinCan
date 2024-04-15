REM  This script adds CMOC Admins to servers that are in a DMZ.
REM  IT will create the account and then add it to the administrators group
REM  There is no option to automate the "force password at next logon"
REM  The initial password for these users will "Password1!"
REM  Written by Kelly Ethington, 573059, RESTON VA   Date:  02/28/2008
REM  *********************************************************************
REM  


net user zz118169 Password1! /fullname:"Robert Ayala" /add
net user zzbusteji Password1! /fullname:"Jimmie Buster" /add
net user zzCHOMAGR Password1! /fullname:"Gregg Choma" /add
net user zzchuibe Password1! /fullname:"Benny Chui" /add
net user zzchunby Password1! /fullname:"Byong J. Chun" /add
net user zzClaytBr Password1! /fullname:"Bryan Clayton" /add
net user zzcruztr Password1! /fullname:"Tracey Cruz" /add
net user zzdewbrry Password1! /fullname:"Les Dewbre" /add
net user zzFARLESC Password1! /fullname:"Scott Farley" /add
net user zzGLUNTJA Password1! /fullname:"James Glunt" /add
net user ZZGRIMDE Password1! /fullname:"Debbie Grim" /add
net user zz108066 Password1! /fullname:"Julio Hernandez" /add
net user zzHillDa2 Password1! /fullname:"David L. Hill" /add
net user zzHonstam Password1! /fullname:"Amy C. Honstetter" /add
net user zzhortich Password1! /fullname:"Christopher R. Horting" /add
net user zzhuntegr Password1! /fullname:"Greg Hunter" /add
net user zzjacksli Password1! /fullname:"Linda L. Jackson-Smith" /add
net user zzkayma Password1! /fullname:"Martin Kay" /add
net user zzlortomi Password1! /fullname:"Michael Lorton" /add
net user zzMcGouMi Password1! /fullname:"Mitchell McGouldrick" /add
net user zzmullesh2 Password1! /fullname:"Sharon Mullen" /add
net user zzphanph Password1! /fullname:"Phong Phan" /add
net user zzRouilJo Password1! /fullname:"Joseph Rouillier III" /add
net user zzsentiad Password1! /fullname:"AJ Sentiger" /add
net user zzsepulda Password1! /fullname:"David Sepulveda" /add
net user zzstaffra Password1! /fullname:"Gene E. Stafford" /add


net localgroup administrators zz118169 /add
net localgroup administrators zzbusteji /add
net localgroup administrators zzCHOMAGR /add
net localgroup administrators zzchuibe /add
net localgroup administrators zzchunby /add
net localgroup administrators zzClaytBr /add
net localgroup administrators zzcruztr /add
net localgroup administrators zzdewbrry /add
net localgroup administrators zzFARLESC /add
net localgroup administrators zzGLUNTJA /add
net localgroup administrators ZZGRIMDE /add
net localgroup administrators zz108066 /add
net localgroup administrators zzHillDa2 /add
net localgroup administrators zzHonstam /add
net localgroup administrators zzhortich /add
net localgroup administrators zzhuntegr /add
net localgroup administrators zzjacksli /add
net localgroup administrators zzkayma /add
net localgroup administrators zzlortomi /add
net localgroup administrators zzMcGouMi /add
net localgroup administrators zzmullesh2 /add
net localgroup administrators zzphanph /add
net localgroup administrators zzRouilJo /add
net localgroup administrators zzsentiad /add
net localgroup administrators zzsepulda /add
net localgroup administrators zzstaffra /add

net localgroup users
pause
net localgroup administrators
pause