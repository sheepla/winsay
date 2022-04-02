NAME := winsay.exe
BIN := bin/

.PHONY: build
build:
	go build -tags forceposix -o $(BIN)/$(NAME)
