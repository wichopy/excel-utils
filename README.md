To set up dev env,run `go get github.com/360EntSecGroup-Skylar/excelize` to load external dep.

To run dev, `go run main.go` with any flags you need

Available flags and their defaults:
-ignoreTopRow (true)
-targetSheet (Sheet1)
-targetDirectory (.)

To build, `go build`

Since this we developed on a Mac, I added an extra script for building on windows.