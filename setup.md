# Setup Guides for Developers

## System information [not a requisite]
- Python version: 3.11.1
- OS: Windows 11 (64-bit operating system, x64-based processor)

## Setting up virtual environment
[Please skip to the **Too complicated?** section if you do not want to go into details.](#too-complicated)

### Creating venv
Open terminal in your preferred location to set up a virtual environment and run the following commands:

```
pip install virtualenv
```
```
python -m venv <venv_name>
```

We will call the `<venv_name>` **mic-drop-results** from now on, since we are creating a venv for this program.

### Installing dependencies
Then, navigate to `.\mic-drop-results\Scripts` and paste the [**requirements.txt**](requirements.txt) file here. Make sure you are now in the `Scripts` folder, run the following command to switch to venv and install all dependencies:

```
.\activate
```
```
pip install -r requirements.txt
```
It will take a few minutes.

### Selecting the virtual environment's interpreter
Restart your IDE. For Visual Studio Code, open **app.py**, look down the bottom-right corner, click on the Python version next to the word *Python*, and select the venv's interpreter (not the global one).

If you could not find the interpreter, you can type in the path to the venv's **python.exe**:

```
.\mic-drop-results\Scripts\python.exe
```

### Bonus: Freezing dependencies into requirements.txt
After installing a new package, make sure you recompile the **requirements.txt** file using the following command:

```
pip freeze > requirements.txt
```

---

## Too complicated?
You could install the dependencies globally if you do not mind. Simply run a single command:

```
pip install -r requirements.txt
```
