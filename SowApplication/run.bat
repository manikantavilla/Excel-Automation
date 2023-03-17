@echo off
REM set java home
set JAVA_HOME=.\jdk-17

REM Read config.properties file
for /f "tokens=1* delims==" %%G in (config.properties) do set "%%G=%%H"

REM set java args
for /f "tokens=*" %%a in ('%ARGS%') do (set "javaArgs=!javaArgs! %%a")

SET mm=%date:~4,2%
SET dd=%date:~7,2%
SET yy=%date:~12,2%
SET hh=%time:~0,2%
SET min=%time:~3,2%
SET ss=%time:~6,2%

REM run application with java args
echo starting application ....
echo "check log_<timeStamp> file for status"
REM echo after 1 minute, Open http://localhost:8080/sowForecast.html in any browser

REM Starting Chrome
echo Strating chrome
start chrome http://localhost:8080/sowForecast.html

"%JAVA_HOME%\bin\java.exe" %javaArgs% -jar Automation-0.0.1-SNAPSHOT.jar > log_%mm%%dd%%yy%_%hh%%min%%ss%.log
