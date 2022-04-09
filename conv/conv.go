package conv

import (
	"encoding/csv"
	"io"

	"github.com/goark/csvdata"
	"github.com/goark/csvdata/exceldata"
	"github.com/goark/errs"
)

func ToCsv(w io.Writer, r *exceldata.Reader, comma rune, winNewline bool) error {
	csvw := csv.NewWriter(w)
	csvw.UseCRLF = winNewline
	if comma != 0 {
		csvw.Comma = comma
	}
	defer csvw.Flush()

	rc := csvdata.NewRows(r, false)
	for {
		if err := rc.Next(); err != nil {
			if errs.Is(err, io.EOF) {
				break
			}
			return errs.Wrap(err)
		}
		if err := csvw.Write(rc.Row()); err != nil {
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
