@echo off
echo Starting local web server for Golden Auto Inventory...
echo Please leave this window open while browsing the site locally.
cd webapp
start http://localhost:8000
python -m http.server 8000
