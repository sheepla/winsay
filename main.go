package main

import (
	"fmt"
	"os"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/jessevdk/go-flags"
)

const (
	appName  = "winsay"
	appUsage = "[OPTIONS] TEXT..."
)

type exitCode int

const (
	exitCodeOK exitCode = iota
	exitCodeErrArgs
	exitCodeErrSay
)

type options struct {
	Rate int `short:"r" long:"rate" description:"Speech rate" default:"0"`
}

func main() {
	os.Exit(int(Main(os.Args[1:])))
}

func Main(cliArgs []string) exitCode {
	var opts options
	parser := flags.NewParser(&opts, flags.Default)
	parser.Name = appName
	parser.Usage = appUsage
	args, err := parser.ParseArgs(cliArgs)
	if err != nil {
		if flags.WroteHelp(err) {
			return exitCodeOK
		} else {
			fmt.Fprintf(os.Stderr, "Parse Error: %s\n", err)
			return exitCodeErrArgs
		}
	}
	if len(args) == 0 {
		fmt.Fprintln(os.Stderr, "Must require arguments")
		return exitCodeErrArgs
	}

	if err = say(strings.Join(args, " "), opts.Rate); err != nil {
		fmt.Fprintln(os.Stderr, err)
		return exitCodeErrSay
	}
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
