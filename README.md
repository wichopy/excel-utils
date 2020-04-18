## Usage

1. Download the release for your operating system
2. Open your command line of choice (Terminal, iTerm, PowerShell, etc)
3. Go to where you downloaded the file in your command line and execute the binary:
```
cd Downloads
./excel-merger

OR

./excel-merger.exe -targetSheet "Some Sheet Name"

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