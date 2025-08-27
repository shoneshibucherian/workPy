@echo off
echo Updating the system...
echo.

git reset --hard HEAD
git clean -xffd
git pull

echo.
echo Update complete.