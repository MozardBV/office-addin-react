import OfficeDocument from "./OfficeDocument";
import axios from "axios";

export default class Middleware {
  static async getDocTypes(userProperties, documentProperties) {
    return new Promise((resolve, reject) => {
      const boundary = `boundary_string_${Date.now().toString()}`;
      axios
        .post(
          `https://officetest.mozard.nl/public/index.php?destdossier=${documentProperties.dossierId}&destomgeving=${userProperties.env}`,
          "",
          {
            headers: {
              Authorization: `MOZTOKEN appcode=${userProperties.auth}`,
              "Content-Type": `multipart/form-data; boundary="${boundary}"`,
            },
          }
        )
        .then((res) => {
          resolve(res);
        })
        .catch((e) => {
          reject(new Error(e));
        });
    });
  }

  static async sendFile(progressCallback, userProperties, documentProperties) {
    OfficeDocument.getFile().then((file) => {
      progressCallback(undefined, `Bezig met ophalen bestand (${parseFloat(file.file.size / 1048576).toFixed(2)} MiB)`);

      const boundary = Date.now().toString();

      const sendSlice = () => {
        OfficeDocument.getSlice(file).then((res) => {
          const slice = OfficeDocument.formatSlice(documentProperties, res, boundary);

          axios
            .post(
              `https://officetest.mozard.nl/public/index.php?destdossier=${documentProperties.dossierId}&destdocnummer=${documentProperties.documentId}&destdoctype=${documentProperties.documentType}&destomgeving=${userProperties.env}`,
              slice.buffer,
              {
                headers: {
                  Authorization: `MOZTOKEN appcode=${userProperties.auth}`,
                  "Content-Type": `multipart/form-data; boundary="------------------------${boundary}"`,
                  "X-Moz-Slice": Number(file.counter),
                  "X-Moz-Slice-Index": Number(file.sliceCount) - 1,
                  "X-Moz-SliceHash": btoa(documentProperties.documentId.concat(userProperties.env)),
                },
              }
            )
            .then((res) => {
              file.counter++;
              if (file.counter < file.sliceCount) {
                // Recursion!
                sendSlice();
              } else {
                OfficeDocument.closeFile(file);
                progressCallback(100, "", "Bestand verzonden!");
              }
            })
            .catch((e) => {
              console.error(e);
              throw new Error(e);
            });
        });
      };

      sendSlice();
    });
  }
}
