@echo off
rem set JAVA_HOME=C:\Program Files\Java\jdk-11.0.16
rem set CLASSPATH=C:Automation-0.0.1-SNAPSHOT.jar

REM Read config.properties file
rem For /F "tokens=1* delims==" %A IN (config.properties) DO (IF "%A"=="file" set file=%B)

REM Create command line arguments for each property
set ARGS=for /f "tokens=1* delims==" %G in (config.properties) do set ARGS=!ARGS! -D%G=%H

REM Run the application
"%JAVA_HOME%\bin\java.exe" %ARGS% -jar Automation-0.0.1-SNAPSHOT.jar



