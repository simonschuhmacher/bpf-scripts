# BPF Scripts

A collection of python scripts that make life with exported list from zkipster a it easier.

Most of the scripts expect an excel file containing registred guests from zkipster. For that, create an export of all the guests from an event. Filtering to particular guest lists will then mostly be done in the scripts directly.

The scripts did work for the 2024 BPF, and will most probably need some adaptions for another edition. Feel free to adapt :)

The script that's maybe most useful is `summary.py`, as this creates a summary of on-site, virtual, workshop registrations etc.

Have fun :)


## Setup

You can use any IDE or editor to edit and run the scripts.
A convenient example is VS Code with the official Python extension.

To run the scripts, create a new venv

```python
python -m venv .venv
```

Then, activate the newly created venv

```python
pip install -r requirements.txt
```

And activate it

```python
.\.venv\Scripts\activate
```

If you're using VS Code with the Python extension installed and have the corresponding directory open in a VS Code window, the virtual environment will be activated automatically in VS Code terminals.
