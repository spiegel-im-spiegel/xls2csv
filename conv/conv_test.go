package conv_test

import (
	"archive/zip"
	"bytes"
	"errors"
	"syscall"
	"testing"

	"github.com/spiegel-im-spiegel/xls2csv/conv"
)

var res = `名前,年齢
Alice,18
Bob,19
太郎,20
花子,21
`

func TestOpenXlsxFileSheet(t *testing.T) {
	testCases := []struct {
		path      string
		pw        string
		sheetName string
		err       error
	}{
		{path: "not-exist.xlsx", pw: "", sheetName: "", err: syscall.ENOENT},
		{path: "testdata/test.xlsx", pw: "", sheetName: "", err: nil},
		{path: "testdata/test.xlsx", pw: "passwd", sheetName: "", err: nil},
		{path: "testdata/test.xlsx", pw: "", sheetName: "Sheet1", err: nil},
		{path: "testdata/test.xlsx", pw: "", sheetName: "Sheet2", err: conv.ErrInvalidSheetName},
		{path: "testdata/test-pw.xlsx", pw: "", sheetName: "", err: zip.ErrFormat},
		{path: "testdata/test-pw.xlsx", pw: "passwd", sheetName: "", err: nil},
	}
	for _, tc := range testCases {
		_, _, err := conv.OpenXlsxFileSheet(tc.path, tc.pw, tc.sheetName)
		if !errors.Is(err, tc.err) {
			t.Errorf("OpenXlsxFileSheet() is \"%+v\", want \"%+v\".", err, tc.err)
		}
	}
}

func TestToCsv(t *testing.T) {
	testCases := []struct {
		path    string
		sheetNo int
		err     error
		csvdata string
	}{
		{path: "testdata/test.xlsx", sheetNo: 0, err: nil, csvdata: res},
		{path: "testdata/test.xlsx", sheetNo: 1, err: conv.ErrInvalidSheetName, csvdata: res},
	}
	for _, tc := range testCases {
		xlsx, _, err := conv.OpenXlsxFileSheet(tc.path, "", "")
		if err != nil && !errors.Is(err, tc.err) {
			t.Errorf("OpenXlsxFileSheet() is \"%+v\", want \"%+v\".", err, tc.err)
		}
		if err == nil {
			buf := &bytes.Buffer{}
			err := conv.ToCsv(buf, xlsx, tc.sheetNo)
			if err != nil && !errors.Is(err, tc.err) {
				t.Errorf("ToCsv() is \"%+v\", want \"%+v\".", err, tc.err)
			}
			str := buf.String()
			if err == nil && str != tc.csvdata {
				t.Errorf("ToCsv() is \"%+v\", want \"%+v\".", str, tc.csvdata)
			}
		}
	}
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
