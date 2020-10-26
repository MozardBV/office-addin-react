/* global Office OfficeExtension Word */

export default class OfficeDocument {
  static getDocumentId() {
    return new OfficeExtension.Promise((resolve, reject) => {
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
    return new OfficeExtension.Promise((resolve, reject) => {
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

  static getFile() {
    // 4 MB slice size
    return new OfficeExtension.Promise((resolve, reject) => {
      Office.context.document.getFileAsync("compressed", { sliceSize: 4194304 }, (result) => {
        // eslint-disable-next-line eqeqeq
        if (result.status == Office.AsyncResultStatus.Succeeded) {
          resolve({
            counter: 0,
            file: result.value,
            sliceCount: result.value.sliceCount,
          });
        } else {
          reject(new Error(result.status));
        }
      });
    });
  }

  static getSlice(state) {
    return new OfficeExtension.Promise((resolve, reject) => {
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
      // Encode de slice data (dat is een byte array) als base64 string
      const u8 = new Uint8Array(slice.data);
      const b64encoded = btoa(String.fromCharCode.apply(null, u8));

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
