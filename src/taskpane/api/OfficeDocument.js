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

/* global Office Word */

import OutlookMailbox from "./OutlookMailbox";
export default class OfficeDocument {
  static getDocumentId() {
    return new Office.Promise((resolve, reject) => {
      if (Office.context.requirements.isSetSupported("WordApi", "1.3")) {
        Word.run(async (context) => {
          const prop = context.document.properties.customProperties.getItemOrNullObject("mzdDocumentId");
          context.load(prop);
          await context.sync();
          if (!prop.isNullObject) {
            resolve(prop);
          } else {
            reject("Geen documentnummer gevonden in dit document");
          }
        });
      } else {
        reject("Vereiste Office API voor OfficeDocument.getDocumentId() niet ondersteund");
      }
    });
  }

  static setDocumentId(documentId) {
    return new Office.Promise((resolve, reject) => {
      if (Office.context.requirements.isSetSupported("WordApi", "1.3")) {
        Word.run(async (context) => {
          context.document.properties.customProperties.add("mzdDocumentId", documentId);
          context
            .sync()
            .then((res) => {
              resolve(res);
            })
            .catch((e) => {
              reject(e);
            });
        });
      } else {
        reject("Vereiste Office API voor OfficeDocument.setDocumentId() niet ondersteund");
      }
    });
  }

  static getFile(platform, sliceSize) {
    return new Office.Promise((resolve, reject) => {
      if (platform !== "Outlook") {
        Office.context.document.getFileAsync("compressed", { sliceSize }, (result) => {
          // eslint-disable-next-line eqeqeq
          if (result.status == Office.AsyncResultStatus.Succeeded) {
            resolve({
              counter: 0,
              file: result.value,
              sliceCount: result.value.sliceCount,
            });
          } else {
            reject(result.status);
          }
        });
      } else {
        OutlookMailbox.getEmail()
          .then((res) => {
            resolve(res);
          })
          .catch((e) => {
            reject(e);
          });
      }
    });
  }

  static getSlice(state) {
    return new Office.Promise((resolve, reject) => {
      state.file.getSliceAsync(state.counter, (result) => {
        // eslint-disable-next-line eqeqeq
        if (result.status == Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(result.status));
        }
      });
    });
  }

  static formatSlice(documentProperties, slice, boundary) {
    if (slice.data) {
      let b64encoded;

      if (typeof slice.data !== "string") {
        const u8 = new Uint8Array(slice.data);
        b64encoded = btoa(String.fromCharCode.apply(null, u8));
      } else {
        b64encoded = btoa(slice.data);
      }

      const attachmentContentType = "application/octet-stream";
      const contentDisposition = `Content-Disposition: form-data; name="file"; filename="${documentProperties.documentName}.${documentProperties.documentExtension}"`;

      const requestBodyBeginning = `--------------------------${boundary}\r\nContent-Type: ${attachmentContentType}\r\n${contentDisposition}\r\n\r\n`;
      const requestBodyEnd = `\r\n--------------------------${boundary}--`;

      // Base64-encoded string decoden
      const byteCharacters = atob(b64encoded);

      // De code point (charCode) van elke character wordt de value van de
      // byte. We maken een array met byte values door .charCodeAt() aan te
      // roepen voor elke character in de string.
      const byteNumbers = new Array(byteCharacters.length);

      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }

      // Omzetten naar een echte byte array met de juiste type. (Oftewel,
      // een array van 8-bit unsigned integers.)
      const byteArray = new Uint8Array(byteNumbers);

      const totalRequestSize = requestBodyBeginning.length + byteArray.byteLength + requestBodyEnd.length;

      const uint8array = new Uint8Array(totalRequestSize);

      {
        let i;

        // Het begin van de request toevoegen
        for (i = 0; i < requestBodyBeginning.length; i++) {
          uint8array[i] = requestBodyBeginning.charCodeAt(i) & 0xff;
        }

        // De binary attachment toevoegen
        for (let j = 0; j < byteArray.byteLength; i++, j++) {
          uint8array[i] = byteArray[j];
        }

        // Het eind van de request toevoegen
        for (let j = 0; j < requestBodyEnd.length; i++, j++) {
          uint8array[i] = requestBodyEnd.charCodeAt(j) & 0xff;
        }
      }

      return uint8array;
    } else {
      throw new Error("Slice bevat geen data");
    }
  }

  static closeFile(state) {
    state.file.closeAsync((result) => {
      if (result.status !== "succeeded") {
        throw new Error("Bestand verzonden, maar kon niet gesloten worden");
      }
    });
  }
}
