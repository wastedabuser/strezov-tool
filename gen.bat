echo off
set PATH="%cd%\perl\bin"
perl -v
perl gen.pl -check
perl gen.pl -h
perl gen.pl -process xls
pause