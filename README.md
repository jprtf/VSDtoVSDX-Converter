# VSD to VSDX Converter
A Powershell utility to mass convert Microsoft Visio VSD files to VSDX.

This script uses the Visio COM object to perform the conversion. Ensure that you have Microsoft Visio installed on the machine where you run this script.

## Options to make it an executable (for convenience)

### Compile
You can compile the script using [PS2EXE](https://github.com/MScholtes/PS2EXE). E.g.:

  `Invoke-PS2EXE .\vsd_converter.ps1 .\vsd_converter.exe -noConsole -version '1.0'`

Note that scripts compiled with PS2EXE are often mistakenly detected as malware. The best way to prevent this is to certify the executable.

### Turn it into a polyglot script
1. Edit the .ps1 script

1. Add the following code to the first line

    `@findstr/v "^@f.*&" "%~f0" | powershell -NoProfile -ExecutionPolicy Bypass -&goto:eof`

1. Save it as as .cmd

More info [here](https://stackoverflow.com/questions/2609985/how-to-run-a-powershell-script-within-a-windows-batch-file).
