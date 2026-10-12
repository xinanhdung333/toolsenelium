@echo off
for /L %%i in (1,1,10) do (
  echo fake %%i %date% %time% >> fake.txt
  git add .
  git commit -m "commit fake %%i - %date%"
  timeout /t 1 >nul
)
git push
echo Da day 10 commit fake xong!
pause