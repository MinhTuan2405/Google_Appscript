SCRIPT_ID = 1a2B3cD4EfGhIjKlMnOpQrStUvWxYz1234567890


push:
	cd EduTrack/source && npx clasp push

clone:
	npx clasp clone $(SCRIPT_ID)

pull:
	cd EduTrack/source && npx clasp pull
