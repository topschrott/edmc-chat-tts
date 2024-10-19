
all: flake8 pylint

flake8:
	flake8 load.py

pylint:
	pylint --recursive=y .

deps.pyz: requirements.txt
	python3 -m pip install edge-tts --target deps
	rm -rf deps/bin
	python3 -m zipapp deps -o deps.pyz -m edge_tts.utils:main

run-dev-docker:
	docker run -ti --rm -v $$PWD:/app -w /app python:3 /bin/bash

prep-dev-docker:
	python3 -m pip install -r requirements.txt
	test -d edmc || git clone https://github.com/EDCD/EDMarketConnector.git edmc
