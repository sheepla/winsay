package main

import (
	"flag"
	"fmt"
	"os"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type exitCode int

const (
	exitCodeOK exitCode = iota
	exitCodeErrArgs
	exitCodeErrSpeak
)

type options struct {
	Rate int
}

func main() {
	os.Exit(int(Main()))
}

func Main() exitCode {
	var opts options
	flag.IntVar(&opts.Rate, "r", 0, "Speech rate (default: 0, slowest :-10, fastest: 10)")
	flag.Parse()

	if len(flag.Args()) == 0 {
		fmt.Fprintln(os.Stderr, "Must require arguments")
		return exitCodeErrArgs
	}

	say(strings.Join(flag.Args(), " "), opts.Rate)
	return exitCodeOK
}

func say(text string, rate int) error {
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("SAPI.SpVoice")
	if err != nil {
		return err
	}

	sapi, err := unknown.QueryInterface(ole.IID_IDispatch)
	defer sapi.Release()
	if err != nil {
		return err
	}

	_, err = oleutil.PutProperty(sapi, "Rate", rate)
	if err != nil {
		return err
	}

	_, err = oleutil.CallMethod(sapi, "Speak", text)
	if err != nil {
		return err
	}

	return nil
}
