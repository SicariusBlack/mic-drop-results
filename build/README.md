## Setting up a virtual environment

### Creating a virtual environment
Open Terminal in a folder of choice and run the following commands:

```
pip install virtualenv
```

Let's say we name our virtual environment **md_results**:
```
python -m venv md_results
```

Here is our directory tree after creating your virtual environment.

> **Preview**
> 
>     .
>     └── md_results
>         ├── Include
>         ├── Lib
>         ├── Scripts
>         └── pyvenv.cfg

### Installing dependencies
Then, navigate to `.\md_results\Scripts` and paste the [**requirements.txt**](build/requirements.txt) file here.

> **Preview**
>
>     .
>     └── md_results
>         ├── Include
>         ├── Lib
>         ├── Scripts
>         │   ├── activate
>         │   ├── activate.bat
>         │   ├── ...
>         │   └── requirements.txt
>         └── pyvenv.cfg

Make sure you are now in the `Scripts` folder, run the following command to switch to venv and install all dependencies:

```
.\activate
```
```
pip install -r requirements.txt
```
It will take a few minutes to install.
