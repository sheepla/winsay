package main

import (
	"flag"
	"log"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type options struct {
	Rate int
}

func main() {
	var opts options
	flag.IntVar(&opts.Rate, "r", 0, "Speech rate (default: 0, slowest :-10, fastest: 10)")
	flag.Parse()
	say(strings.Join(flag.Args(), " "), opts.Rate)
}

func say(text string, rate int) {
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

	oleutil.PutProperty(sapi, "Rate", rate)
	oleutil.CallMethod(sapi, "Speak", text)
}
