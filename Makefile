

VENV=./venv/bin/activate

test: $(VENV)
	. $(VENV) && python3 setup.py test

make_venv: $(VENV)

$(VENV): requirements.txt setup.py
	python3 -m venv venv && . $(VENV) && pip3 install -r requirements.txt

clean:
	rm -rf $(VENV)
