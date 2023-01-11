# Open with Google Docs
This is a small program designed to open document files (doc(x), xls(x), and ppt(x) types) in Google Docs, Spreadsheets and Slides. Designed to work on Windows only, but could be compiled to other platforms as well.

### Setup
Place the executable in a place where you wont delete it. AppData is generally a safe place. So, for setup follow these setups:
- Press `Windows Key + R` to open "Run"
- Type `powershell.exe -command "mkdir $env:APPDATA\OpenWGDocs; (New-Object System.Net.WebClient).DownloadFile('https://github.com/FKLC/OpenWithGoogleDocs/releases/latest/download/OpenWGDocs.exe',\"$env:APPDATA\OpenWGDocs.exe\"); explorer $env:APPDATA\OpenWGDocs"`
- Find `OpenWGDocs.exe` in the folder opened
- Run that file as admin. To do so right-click on the file then click to Run as Administrator.
- This should be all. You can now right-click to document files and use Open with option to open them in Google Docs.

### Usage
When you encounter a document file, right-click and use Open With, then choose `OpenWGDocs.exe`. It will ask you to authenticate your account with Open with Google Docs, accept it, then the program will upload the file and open it in your browser.