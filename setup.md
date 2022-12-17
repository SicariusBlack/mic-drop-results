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

We will call the `<venv_name>` **md_results** from now on.

Here is the directory tree after creating your venv.

    .
    └── md_results
        ├── Include
        ├── Lib
        ├── Scripts
        └── pyvenv.cfg

### Installing dependencies
Then, navigate to `.\md_results\Scripts` and paste the [**requirements.txt**](requirements.txt) file here.

    .
    └── md_results
        ├── Include
        ├── Lib
        ├── Scripts
        │   ├── activate
        │   ├── activate.bat
        │   ├── ...
        │   └── requirements.txt
        └── pyvenv.cfg

Make sure you are now in the `Scripts` folder, run the following command to switch to venv and install all dependencies:

```
.\activate
```
```
pip install -r requirements.txt
```
It will take a few minutes to install.

### Selecting the virtual environment's interpreter
Restart your IDE. For Visual Studio Code, open **app.py**, look down the bottom-right corner, and click on the Python version next to the word *Python*.

<div align="center">

![image](https://user-images.githubusercontent.com/106049382/208242818-7f479c5c-5a61-4025-ba6a-ec9d888069c6.png)
</div>

Then, select the venv's interpreter.

<div align="center">

![image](https://user-images.githubusercontent.com/106049382/208242370-46ac666d-ab59-4b2f-a688-ff3c27a8d28c.png)
</div>

If you could not find the interpreter, you can type in the path to the venv's **python.exe**:

```
.\md_results\Scripts\python.exe
```

**You are all set! Happy contributing. I hope you'll have a lot of fun playing around with the code just as much as I do.**

---

## Too complicated?
You could install the dependencies globally if you do not mind. Simply run a single command:

```
pip install -r requirements.txt
```
