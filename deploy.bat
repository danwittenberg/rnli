@echo off
echo Shipping backend to Google...
call clasp push
echo Shipping frontend to GitHub...
git add .
git commit -m "automated deployment"
git push origin main
echo All done!