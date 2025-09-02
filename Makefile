SCRIPT_ID = 1a2B3cD4EfGhIjKlMnOpQrStUvWxYz1234567890

push:
	cd EduTrack/source && npx clasp push

clone:
	npx clasp clone $(SCRIPT_ID)

pull:
	cd EduTrack/source && npx clasp pull

sync:
	make pull
	git add .
	git commit -m "sync at $(shell powershell -Command "Get-Date -Format 'yyyy-MM-dd HH:mm:ss'")"
	git push -u origin main

sync-c:
	make pull
	git add .
	git commit -m "sync at $(shell echo %DATE% %TIME%)"
	git push -u origin main
