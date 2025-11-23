## pre 

winget install git.git Github.cli
winget install -e  AutoHotkey.AutoHotkey -v 1.1.37.02
winget pin add --id AutoHotkey.AutoHotkey --blocking

	python kanakanji.py gemini.txt -o test.txt --debug --log log.file
