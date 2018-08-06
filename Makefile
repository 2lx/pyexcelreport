init:
	pip install -r requirements.txt

test:
	python test_sample.py

.PHONY: init test