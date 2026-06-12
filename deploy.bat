@echo off
echo Shipping backend to Google...
call clasp push

@echo off
echo Pushing latest code to Google...
call clasp push

echo Deploying new web app version...
call clasp deploy -i AKfycbzm5vM3xXhxAkR6Gk7gN-dG3ynZZjGvrqe4Ewqlqy6k -d "Automated Sortable Date Update"

echo All done!
pause
echo Shipping frontend to GitHub...
git add .
git commit -m "automated deployment"
git push origin main
echo All done!