@echo off
cd %1
c:\jdk1.4\bin\javap -c %2 > disasm.txt