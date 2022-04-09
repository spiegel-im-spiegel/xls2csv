package conv

import (
	"github.com/goark/csvdata/exceldata"
	"github.com/goark/errs"
)

func OpenXlsxFileSheet(path, password, sheetName string) (*exceldata.Reader, error) {
	xlsx, err := exceldata.OpenFile(path, password)
	if err != nil {
		return nil, errs.Wrap(err, errs.WithContext("path", path), errs.WithContext("sheetName", sheetName))
	}
	r, err := exceldata.New(xlsx, sheetName)
	return r, errs.Wrap(err, errs.WithContext("path", path), errs.WithContext("sheetName", sheetName))
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
