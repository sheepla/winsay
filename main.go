package main

import (
	"log"
	"os"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func main() {
	say(os.Args[1])
}

func say(text string) {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("SAPI.SpVoice")
	if err != nil {
		log.Fatal(err)
	}

	sapi, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Fatal(err)
	}
	defer sapi.Release()
	oleutil.CallMethod(sapi, "Speak", text)
}
