## Usage

1. Download the release for your operating system
2. Open your command line of choice (Terminal, iTerm, PowerShell, etc)
3. Go to where you downloaded the file in your command line and execute the binary:
```
cd Downloads
C:\Users\William Chou\Desktop>  ./excel-merger.exe -targetDirectory "C:\Users\William Chou\Downloads" -targetSheet "Some-data" -ignoreTopRow=false
```

Starting excel merger with flags: ignoreTopRow: false, targetSheet: Some-data, targetDirectory:C:\Users\William Chou\Downloads

Scanning files in target directory: 'C:\Users\William Chou\Downloads'

Opening Book1 (1).xlsx
Copying data from sheet Some-data to output.
Opening Book1.xlsx
Copying data from sheet Some-data to output.
Opening output.xlsx
Cannot find target sheet 'Some-data' in file 'output.xlsx'

1210 rows copied

Saving to C:\Users\William Chou\Desktop/output.xlsx

```
PS C:\Users\William Chou\Desktop>  ./excel-merger.exe -targetDirectory "C:\Users\William Chou\Downloads" -targetSheet "Some-data" -ignoreTopRow=true
```

Starting excel merger with flags: ignoreTopRow: true, targetSheet: Some-data, targetDirectory:C:\Users\William Chou\Downloads

Scanning files in target directory: 'C:\Users\William Chou\Downloads'

Opening Book1 (1).xlsx
Copying data from sheet Some-data to output.
Opening Book1.xlsx
Copying data from sheet Some-data to output.
Opening output.xlsx
Cannot find target sheet 'Some-data' in file 'output.xlsx'

1208 rows copied

Saving to C:\Users\William Chou\Desktop/output.xlsx
```

## Dev

To set up dev env,run `go get github.com/360EntSecGroup-Skylar/excelize` to load external dep.

To run dev, `go run main.go` with any flags you need

Available flags and their defaults:
-ignoreTopRow (true)
-targetSheet (Sheet1)
-targetDirectory (.)

## Build

To build, `go build`

Since this we developed on a Mac, I added an extra script for building on windows.
