Write-Host ""
Write-Host "PYTHON INSTALLATION REQUIRED!"
Write-Host ""
Write-Host "PLEASE CONNECT TO THE INTERNET FOR FIRST TIME INSTALLATION."
Write-Host ""
Write-Host ""
$env:FLASK_APP = "KGSweb"
$env:FLASK_ENV = "development"
pip install -e .
python open_web.pyw
flask run
Read-Host -Prompt "Press Enter to exit"