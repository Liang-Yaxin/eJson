<div align="center">

Manage your `json` data better with visual Excel sheets

![npm](https://img.shields.io/badge/npm-v1.0.0-blue)

English | [简体中文](./READEME_Chinese.md)

</div>

## Effect preview

Template Excel file:
![template Excel file ](./media/template_excel.png)

Output json data:
![output json data ](./media/effects.png)

## Introduction
> ` ex-json-cli `, which is composed of three words  'excel',  'json' and  'cli'.

### Practical problems
Sometimes, in order to save the back-end development cost, we will store some data directly in json, and then use these data for rendering. But:

- If there are more pages or modules that need data, the corresponding json files will also increase. There are too many files, the data is too fragmented, and managing so many files becomes troublesome.
- When the amount of data is large, the data will become very lengthy, and it is very difficult to find a specific piece of data directly.
- Sometimes' json' is not a good way to spread our data when we want to share it with other non-developers.

### Why choose Excel to manage data?
- More convenient operation. Excel is a professional office software. It is much more convenient to add, delete and check data in Excel than to modify it in our editor.
- More visual. Each sheet can be regarded as the data of a page or a module; The name of each sheet is the name of our json; Every row of data in the table is the data of every item in our json.
- More convenient management. The data are all integrated in an Excel file, which is more convenient for us to manage the data of all modules or pages.
- More suitable for communication. Json is not suitable for spreading among non-developers, but Excel is suitable for everyone. Not only can non-developers modify this Excel, but if you find something wrong in Excel, you can also modify it, and then synchronize it with others.

### Tools are born
Therefore, in order to solve the above problems and combine the comprehensive advantages of Excel, `ex-JSON-CLI` came into being, so you just need to focus on managing the Excel file. s

## Quick start
### 1. Installation tool
````npm
npm i ex-json-cli -g
````
### 2. Get the template excel file
You don't need to make an excel form yourself, here is a template excel file, you just need to execute:
````npm
ex-json-cli gt
````
Or save the template file to a directory:
````npm
ex-json-cli gt './xlsx_template/'
````
In this way, we can get an excel template file, and then change the data in it into what you want. Yes, it's as simple as that.
### 3. Output the json file.
> you can go to the [Explanation](#Explanation) module to get a better understanding of `-k` and `-s`.
````
ex-json-cli -i './xlsx_template/template.xlsx' -s 3 -r all
````

## Notice
In the process of use, here are the following precautions:
-When generating json data, please ensure that the value of `-s` is correct. `-s` means that data is read from the first row of each ` sheet `, and the default value is  line 3.

## Options and commands
````npm
Usage: ex-json-cli [options] [command]

=> Manage your json data better with visual excel sheets

Options:
-v View current version
-i, --input [path] Path of excel to be converted
-s, --start-row [number] Read data from what row of sheet
-r, --range [string] Select the table you want to read
-h, --help View help

Commands:
gt [path] Get the excel template file
````

### option
| Parameter | Is it necessary | Default value | Description | Supplement
| ---| --- | --- | --- | ---
| `-v` | No || View the current version |
| `-i` | is ||| The path of the excel table to be converted |
| `-s` | No | `3` | Which row in the excel table should read data from | The data is read from the third row by default.
| `-r` | No | `all` | Read that sheet from the excel table | Read all sheets by default.
| `-h` | No || View Help |

### command
| Statement | Parameter | Description
| ---| --- | --- |
|` gt`|` path`|` path` Save the directory for the obtained template excel file. < /br > when path is empty, the obtained template excel file is saved in the xlsx_template folder of the current directory by default.

## <a id="Explanation">Explanation</a>

The overall structure of a table is generally divided into three parts (`top`, middle, bottom):
-The first block, which we call (`t`), refers to the headline of the whole table.
-The second block, which we call (`m`), is an overview of each column of information in the table.
-The third block, which we call (` b`), is the number of rows where the program will start reading data (`-s 3`).

## How to use your own excel sheet
Because everyone's excel is different, we strongly recommend that you use our template excel file (template.xlsx) to manage data, but the style of this file may not satisfy you.

So if you want to 'DIY' the style of the table, according to the analysis of the [Explanation](#Explanation) module, 'DIY' your table needs to follow the following rules:
-modules' t' and' m' are not necessary, * * However, please ensure that the data in your excel file has the same structure as the module' b' in the above figure * *
-Please make sure that your `-s` value is correct.