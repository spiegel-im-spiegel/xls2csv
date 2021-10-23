package facade

import (
	"os"
	"runtime"

	"github.com/spf13/cobra"
	"github.com/spiegel-im-spiegel/errs"
	"github.com/spiegel-im-spiegel/gocli/exitcode"
	"github.com/spiegel-im-spiegel/gocli/rwi"
	"github.com/spiegel-im-spiegel/xls2csv/conv"
)

var (
	//Name is applicatin name
	Name = "xls2csv"
	//Version is version for applicatin
	Version = "dev-version"
)

var (
	versionFlag bool //version flag
	debugFlag   bool //debug flag
)

//newRootCmd returns cobra.Command instance for root command
func newRootCmd(ui *rwi.RWI, args []string) *cobra.Command {
	rootCmd := &cobra.Command{
		Use:   Name + " [flags] <Excel file>",
		Short: "Export CSV text from Excel data",
		Long:  "Export CSV text from Excel data.",
		RunE: func(cmd *cobra.Command, args []string) error {
			//output version
			if versionFlag {
				return debugPrint(ui, errs.Wrap(ui.OutputErrln(getVersion())))
			}

			//options
			pw, err := cmd.Flags().GetString("password")
			if err != nil {
				return debugPrint(ui, errs.New("Error in --password option", errs.WithCause(err)))
			}
			sheetName, err := cmd.Flags().GetString("sheet")
			if err != nil {
				return debugPrint(ui, errs.New("Error in --sheet option", errs.WithCause(err)))
			}
			out, err := cmd.Flags().GetString("output")
			if err != nil {
				return debugPrint(ui, errs.New("Error in --output option", errs.WithCause(err)))
			}
			tsvFlag, err := cmd.Flags().GetBool("tsv")
			if err != nil {
				return debugPrint(ui, errs.New("Error in --tsv option", errs.WithCause(err)))
			}
			winNewline, err := cmd.Flags().GetBool("win-newline")
			if err != nil {
				return debugPrint(ui, errs.New("Error win-newline --tsv option", errs.WithCause(err)))
			}

			// output stream
			w := ui.Writer()
			if len(out) > 0 {
				file, err := os.Create(out)
				if err != nil {
					return debugPrint(ui, errs.Wrap(err, errs.WithContext("output", out)))
				}
				defer file.Close()
				w = file
			}

			// open Excel file
			if len(args) != 1 {
				return debugPrint(ui, errs.Wrap(os.ErrInvalid, errs.WithContext("args", args)))
			}
			xlsx, err := conv.OpenXlsxFileSheet(args[0], pw, sheetName)
			if err != nil {
				return debugPrint(ui, errs.Wrap(err, errs.WithContext("path", args[0]), errs.WithContext("password", pw), errs.WithContext("sheet", sheetName)))
			}

			// export CSV
			comma := rune(0)
			if tsvFlag {
				comma = '\t'
			}
			return debugPrint(ui, conv.ToCsv(w, xlsx, comma, winNewline))
		},
	}
	rootCmd.Flags().BoolVarP(&versionFlag, "version", "v", false, "output version of "+Name)
	rootCmd.Flags().BoolVarP(&debugFlag, "debug", "", false, "for debug")
	rootCmd.Flags().StringP("sheet", "s", "", "sheet name in Excel file")
	rootCmd.Flags().StringP("password", "p", "", "password in Excel file")
	rootCmd.Flags().StringP("output", "o", "", "path of output CSV file")
	rootCmd.Flags().BoolP("tsv", "t", false, "output with TSV format")
	rootCmd.Flags().BoolP("win-newline", "w", false, "output with CRLF newline")

	rootCmd.SilenceUsage = true
	rootCmd.SetArgs(args)
	rootCmd.SetIn(ui.Reader())       //Stdin
	rootCmd.SetOut(ui.ErrorWriter()) //Stdout -> Stderr
	rootCmd.SetErr(ui.ErrorWriter()) //Stderr

	return rootCmd
}

//Execute is called from main function
func Execute(ui *rwi.RWI, args []string) (exit exitcode.ExitCode) {
	defer func() {
		//panic hundling
		if r := recover(); r != nil {
			_ = ui.OutputErrln("Panic:", r)
			for depth := 0; ; depth++ {
				pc, src, line, ok := runtime.Caller(depth)
				if !ok {
					break
				}
				_ = ui.OutputErrln(" ->", depth, ":", runtime.FuncForPC(pc).Name(), ":", src, ":", line)
			}
			exit = exitcode.Abnormal
		}
	}()

	//execution
	exit = exitcode.Normal
	if err := newRootCmd(ui, args).Execute(); err != nil {
		exit = exitcode.Abnormal
	}
	return
}

/* Copyright 2021 Spiegel
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * 	http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
