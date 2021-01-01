# Trello JSON to Execl

## What Is This?

This is a simple Python application intended to provide a Excel file from a
JSON trello file given. It extracts certain information from the cards,
which respects a very specific format.
It is made to facilitate the calculation of ETA on a project.

### Format accepted

Make sure your card titles respect the following format :

```txt
[Title] #SuenceNumber 'Discription' (ETA h)
```

## Installation

Make sure to have Python 3.8.3 or above
Install this libraries by running

```sh
pip install simplejson
pip install re
pip installpandas
pip installtkinter
pip installxlsxwriter
```

## How To Use This

Douple click on app.py or run

```python
py app.py
```

## Development

If you want to work on this application we’d love your pull requests and tickets on GitHub!

1. If you open up a ticket, please make sure it describes the problem or feature request fully.
2. If you send us a pull request, make sure you add a test for what you added, and make sure the full test suite runs with `make test`.
