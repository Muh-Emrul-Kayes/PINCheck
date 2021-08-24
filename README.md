
## Inputs

* A [Library Path](Work_Dir/ip_lib/) Conating:
  * LEF File
  * LIB File
  * Verilog File

* [Output Directory](Work_Dir/sample_op_dir)


## Outputs

* Excel Report File


## Getting started

### Step 1: Open powershell and Clone the PINCheck source code to your local environment
```console
$ git clone https://github.com/Muh-Emrul-Kayes/PINCheck
$ cd .\PINCheck\Work_Dir\
```

### Step 2: Create a [Python virtualenv](https://docs.python.org/3/tutorial/venv.html)
Note: Please DO NOT skip this step!
```console
$ python -m virtualenv general
$ .\general\Scripts\activate.ps1
```

### Step 3: Install requirement.txt File
```console
$ pip install -r ./requirement.txt
```


### Step 4: Run Project
```console
$ python .\PinCheck.py
```

## Deactivate Python virtualenv
```
$ deactivate
```
