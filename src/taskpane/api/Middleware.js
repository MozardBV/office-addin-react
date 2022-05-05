// office-addin-react - Koppeling van Mozard met Microsoft Office
// Copyright (C) 2021  Mozard BV
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.

import OfficeDocument from "./OfficeDocument";
import axios from "axios";

export default class Middleware {
  static async getDocTypes(userProperties, documentProperties) {
    return new Promise((resolve, reject) => {
      const boundary = `boundary_string_${Date.now().toString()}`;
      axios
        .post(`/public/index.php?destdossier=${documentProperties.dossierId}&destomgeving=${userProperties.env}`, "", {
          headers: {
            Authorization: `MOZTOKEN appcode=${userProperties.auth}`,
            "Content-Type": `multipart/form-data; boundary="${boundary}"`,
          },
        })
        .then((res) => {
          resolve(res);
        })
        .catch((e) => {
          reject(new Error(e));
        });
    });
  }

  static async sendFile(progressCallback, errorCallback, userProperties, documentProperties) {
    OfficeDocument.getFile(userProperties.platform, 65536)
      .then((file) => {
        progressCallback(
          undefined,
          `Bezig met ophalen bestand (${parseFloat(file.file.size / 1048576).toFixed(2)} MiB)`
        );

        const boundary = Date.now().toString();

        const submit = (data) => {
          progressCallback(parseInt((file.counter / (file.sliceCount - 1)) * 100), "", "Bezig met verzenden");

          axios
            .post(
              `/public/index.php?destdossier=${documentProperties.dossierId}&destdocnummer=${documentProperties.documentId}&destdoctype=${documentProperties.documentType}&destomgeving=${userProperties.env}`,
              data,
              {
                headers: {
                  Authorization: `MOZTOKEN appcode=${userProperties.auth}`,
                  "Content-Type": `multipart/form-data; boundary="------------------------${boundary}"`,
                  "X-Moz-Slice": Number(file.counter),
                  "X-Moz-Slice-Index": Number(file.sliceCount) - 1,
                  "X-Moz-SliceHash": btoa(documentProperties.documentId + userProperties.env),
                },
              }
            )
            .then((res) => {
              console.log(res);
              file.counter++;
              if (file.counter < file.sliceCount) {
                // Recursion!
                // Outlook komt hier nooit, omdat die altijd maar 1 slice heeft.
                // eslint-disable-next-line @typescript-eslint/no-use-before-define
                sendSlice();
              } else {
                if (userProperties.platform !== "Outlook") OfficeDocument.closeFile(file);
                progressCallback(100, "", "Bestand verzonden!");
              }
            })
            .catch((e) => {
              errorCallback(e);
              console.error(e);
              throw new Error(e);
            });
        };

        const sendSlice = () => {
          OfficeDocument.getSlice(file).then((res) => {
            const slice = OfficeDocument.formatSlice(documentProperties, res, boundary);

            submit(slice.buffer);
          });
        };

        if (userProperties.platform !== "Outlook") {
          sendSlice();
        } else {
          const email = OfficeDocument.formatSlice(
            documentProperties,
            {
              data: file.file,
            },
            boundary
          );

          submit(email);
        }
      })
      .catch((e) => {
        errorCallback(e);
        console.error(e);
      });
  }
}
