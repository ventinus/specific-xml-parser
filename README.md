# Tallgrass parser

- open command prompt
- navigate to the package source code (cd Desktop/"tallgrass parser")
- save the .xlsx file to this directory, rename it to "file" or whatever is short and easy
- to simply execute the script with default arguments, type: `$ ruby parser.rb`
- files will be generated in the /output directory with each file the name of the sheet it parsed

### Arguments
- `--start_row`: Number, the first row where data begins, not including the delivery row. Default value is 10
- `--sheet`: Number, specifies the sheet to parse. Default is to parse all sheets.
- `--filename`: String, the name of the file (without extension) if file was saved under a different name. Default is file
- `--show-empty`: Boolean, if provided, the output will include a note for the rows where the value is 0. Default value is false

### Examples
```
$ ruby parser.rb --filename custom-filename

$ ruby parser.rb --sheet 3 --sheet 4 --sheet 7

$ ruby parser.rb --start_row 12

$ ruby parser.rb --start_row 12 --show-empty true
```
