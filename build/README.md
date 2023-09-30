# Builder's Guide

## 1. Setting up a virtual environment

### Creating a virtual environment
Open Terminal in a folder of choice and run the following command:

```
pip install virtualenv
```

Let's say we name our virtual environment **md_results**:
```
python -m venv md_results
```

Here is our directory tree after creating your virtual environment.

> ##### Preview
> 
>     ▼
>     └── ▼ md_results
>         ├── ▶ Include
>         ├── ▶ Lib
>         ├── ▶ Scripts
>         └── pyvenv.cfg

### Installing dependencies
After creating the virtual environment, navigate to `.\md_results\Scripts`, download [**requirements.txt**](requirements.txt), and paste it there.

> ##### Preview
>
>     ▼
>     └── ▼ md_results
>         ├── ▶ Include
>         ├── ▶ Lib
>         ├── ▼ Scripts
>         │   ├── activate
>         │   ├── activate.bat
>         │   ├── ...
>         │   └── requirements.txt
>         └── pyvenv.cfg

Make sure you are in the `Scripts` folder, run the following commands to activate the virtual environment and install all dependencies:

```
.\activate
```
```
pip install -r requirements.txt
```

> [!IMPORTANT]  
> Test

It will take a few minutes to install.
