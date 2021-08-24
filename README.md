
## Inputs

* A [Library Path](WorkDir/ip_lib/) Conating:
  * LEF File
  * LIB File
  * Verilog File

* [Output Directory](WorkDir/sample_op_dir)


## Outputs

* Excel Report File


## Getting started

### Step 1: Open powershell and Clone the PINCheck source code to your local environment
```console
$ git clone https://github.com/
$ cd Work_Dir
```

### Step 2: Create a [Python virtualenv](https://docs.python.org/3/tutorial/venv.html)
Note: You may choose to skip this step if you are doing a system-wide install for multiple users.
      Please DO NOT skip this step!
```console
$ python -m virtualenv general
$ general/bin/activate
```

### Step 3a: Install requirement.txt File
```console
$ pip install -r ./requirement.txt
```


### Step 4: Run Project
```console
$ python .\PinCheck.py"
```