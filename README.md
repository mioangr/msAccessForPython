# msAccessForPython
Python code wrapper of pyodbc to read and write Microsoft Access databases
To install it in your environment as a package called "msAccessRW", run:

```powershell
conda activate your-environment-name
python -m pip install -e .
```

The module is installed into the currently active conda environment. Replace
`your-environment-name` with the environment where you want to use it.

Alternatively, specify the conda environment without activating it:

```powershell
conda run -n your-environment-name python -m pip install -e .
```

The `-e` option keeps the package linked to this source folder, so source
changes are used the next time the module runs.


The same commands work in regular Windows Command Prompt (cmd.exe):

```
conda activate your-environment-name
python -m pip install -e .
```

Or without activating the environment:

```
conda run -n your-environment-name pip install -e .
```

The syntax is the same. However, conda must be initialized and available in cmd.exe, usually by opening Anaconda Prompt or Miniconda Prompt.