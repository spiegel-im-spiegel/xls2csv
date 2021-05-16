package conv

import (
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/spiegel-im-spiegel/errs"
)

func OpenXlsxFileSheet(path, password, sheetName string) (*excelize.File, int, error) {
	xlsx, err := excelize.OpenFile(path, excelize.Options{Password: password})
	if err != nil {
		return xlsx, 0, errs.Wrap(err, errs.WithContext("path", path), errs.WithContext("sheetName", sheetName))
	}
	sheetNo := 0
	if len(sheetName) > 0 {
		sheetNo := xlsx.GetSheetIndex(sheetName)
		if sheetNo < 0 {
			return xlsx, -1, errs.Wrap(ErrInvalidSheetName, errs.WithContext("path", path), errs.WithContext("sheetName", sheetName))
		}
	}
	return xlsx, sheetNo, nil
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
