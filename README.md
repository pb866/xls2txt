Julia package xls2txt
=====================

Convert FlightAware online flight data from Excel to text files.


Installation
------------

Install into your preferred environment with

```julia
julia> ]
pkg> activate <ENV>
pkg> add https://github.com/pb866/xls2txt.git
```


Using xls2txt
-------------

In order to convert Excel files into textfiles, you need to save all your Excel files
in subfolders for each flight route and save those subfolders in a folder.
Activate the environment with your `xls2txt` package and

```julia
import xls2txt
```

Now you only need to define a logger for the output warnings of unconverted files
in `Warnings.log` and call the function correctXLS with the folder paths of your
input folder and a folder paths of the output folder, where `.dat` files will be
saved with the same file names as the Excel files in the same subfolders as in the
input folder. Folder paths can be relative or absolute.

Als Excel sheets must have the same name. The default is the German standard name
`"Tabelle1"`. If you need to change that name, use an optional third input argument
in `correctXLS` for the sheet name.

E.g., run the following code in the REPL to convert Excel files in `data/in`
to text files in `data/out` of Excel files with the sheet name `sheet1`.

```julia
logger = logg.FileLogger("Warnings.log")
logg.global_logger(logger)
correctXLS("data/in", "data/out/", "sheet1")
```

Alternatively, you can save the following code as `Julia` script

```julia
#! /usr/local/bin/julia

import Pkg; Pkg.activate("xls2txt")
using xls2txt

logger = logg.FileLogger("Warnings.log")
logg.global_logger(logger)
correctXLS(ARGS...)
```

and then run `Julia` from console with

```julia
./<script name> <input folder> <output folder> [<sheet name>]
```


Results
-------

The script will write to your specified output folder. If it doesn't exist, it will
be created. In your folders the same subfolders in in the specified input folder
will be created and within `tab`-separated text (`.dat`) files with the same content
and file names as the `Excel` files.
