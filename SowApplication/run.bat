@echo off
set JAVA_HOME=C:\Program Files\Java\jdk-17
set CLASSPATH=C:\Automation-0.0.1-SNAPSHOT.jar;C:\config.properties

REM Read config.properties file
for /f "tokens=1* delims==" %%G in (config.properties) do set "%%G=%%H"

REM Create command line arguments for each property
set ARGS=for /f "tokens=1* delims==" %%G in (config.properties) do set ARGS=!ARGS! -D%%G=%%H

REM Run the application
for /f "tokens=*" %%a in ('%ARGS%') do (set "javaArgs=!javaArgs! %%a")

"%JAVA_HOME%\bin\java.exe" %javaArgs% -jar Automation-0.0.1-SNAPSHOT.jar
