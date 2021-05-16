package conv

import (
	"encoding/csv"
	"io"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/spiegel-im-spiegel/errs"
)

func ToCsv(w io.Writer, xlsx *excelize.File, sheetIndex int, comma rune, winNewline bool) error {
	rows, err := xlsx.Rows(xlsx.GetSheetName(sheetIndex))
	if err != nil {
		var errSheet excelize.ErrSheetNotExist
		if errs.As(err, &errSheet) {
			return errs.Wrap(ErrInvalidSheetName, errs.WithCause(err))
		}
		return errs.Wrap(err)
	}
	csvw := csv.NewWriter(w)
	csvw.UseCRLF = winNewline
	if comma != 0 {
		csvw.Comma = comma
	}
	defer csvw.Flush()
	for rows.Next() {
		cols, err := rows.Columns()
		if err != nil {
			return errs.Wrap(err)
		}
		if err := csvw.Write(cols); err != nil {
			return errs.Wrap(err)
		}
	}
	return nil
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
