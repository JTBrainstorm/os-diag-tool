@echo off
rem Batch file to export registry. Registry path and Output path must be specified as input arguments
set RegistryPath=%1
set OutputPath=%2
reg export %RegistryPath% %OutputPath%
