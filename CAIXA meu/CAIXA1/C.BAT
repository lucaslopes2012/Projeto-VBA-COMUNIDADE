echo off
\clipper\clipper %1
\clipper\tlink %1 \clipper\ndx \clipper\debug,,,\clipper\extend \clipper\clipper/x
del %1.obj
echo Digite ^C para cancelar a execucao ou
pause
echo on
%1
