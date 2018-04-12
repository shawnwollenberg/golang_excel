# golang_excel
whips through multiple excel files and combines them into one consolidated file

to build exe 32 and 64 bit...
export GOOS=windows
export GOARCH=386
export CGO_ENABLED=1
export CC=i586-mingw32msvc-gcc
go build -o ${PWD##*/}_32bit.exe -ldflags "-extldflags -static"

export GOOS=windows
export GOARCH=amd64
export CGO_ENABLED=1
export CC=i586-mingw32msvc-gcc
go build -o ${PWD##*/}_64bit.exe -ldflags "-extldflags -static"
