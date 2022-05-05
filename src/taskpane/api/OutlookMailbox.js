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

/* global Office */

import { fileTypeFromBuffer } from "file-type";

export default class OutlookMailbox {
  static getSubject() {
    // Filter disallowed characters
    return Office.context.mailbox.item.subject;
  }

  static getAttachments(attachments) {
    return new Office.Promise((resolve, reject) => {
      const getAttachmentAsync = async (attachment) => {
        const { id, name, size } = attachment;
        return this.getAttachment(id, name, size);
      };

      const getAllAttachments = async () => {
        return Promise.all(attachments.map((attachment) => getAttachmentAsync(attachment)));
      };

      getAllAttachments()
        .then((result) => {
          const filtered = result.filter(Boolean);
          resolve(filtered);
        })
        .catch((e) => {
          reject(e);
        });
    });
  }

  static getEmail() {
    return new Office.Promise((resolve, reject) => {
      if (Office.context.requirements.isSetSupported("Mailbox", "1.8")) {
        const boundary = Date.now().toString();
        Office.context.mailbox.item.getAllInternetHeadersAsync(async (headers) => {
          // .getAllInternetHeadersAsync geeft soms een 5001 error op Mac OS of succeeded zonder headers
          // Zie: https://github.com/OfficeDev/office-js/issues/2386
          if (headers.status === "failed" || !headers.value) {
            reject(
              headers.error || {
                name: "OfficeAPI interne fout",
                message: "Er is een interne fout opgetreden bij de OfficeAPI",
                code: 5001,
              }
            );
            return;
          }

          // Content-Type eraf halen, en er zelf eentje zetten.
          const arr = headers.value.split("\n");
          const indexContentType = arr.findIndex((el) => el.toLowerCase().startsWith("content-type"));
          const indexNextHeader = arr.findIndex((el, idx) => idx > indexContentType && el.includes(":"));
          const arr1 = arr.slice(0, indexContentType);
          const arr2 = arr.slice(indexNextHeader);
          headers = arr1.concat(arr2).join("\n");
          headers += `Content-type: multipart/alternative; boundary="${boundary}"\r\n`;
          headers += "\r\n";

          const attachments = Office.context.mailbox.item.attachments;
          const hasAttachments = attachments.length > 0;
          const attachmentContents = hasAttachments ? await this.getAttachments() : undefined;

          Promise.all([this.getEmailAsText(boundary), this.getEmailAsHTML(boundary)])
            .then((res) => {
              const mainBody = res.join("");
              let eml = headers.concat(mainBody);
              if (hasAttachments) {
                const attachmentsArr = attachmentContents.map((attachment) => {
                  const { headers, body } = attachment;
                  let attachmentHeaders = `--${boundary}\r\n`;
                  attachmentHeaders += headers.join("\r\n");
                  attachmentHeaders += "\r\n\r\n";
                  return attachmentHeaders.concat(body);
                });
                const allAttachments = attachmentsArr.join("\r\n");
                eml = eml.concat("\r\n", allAttachments);
              }
              eml = eml.concat(`\r\n\r\n--${boundary}--`);
              resolve({
                counter: 0,
                file: eml,
                sliceCount: 1,
              });
            })
            .catch((e) => {
              reject(e);
            });
        });
      } else {
        reject("Vereiste Office API voor OutlookMailbox.getEmail() niet ondersteund");
      }
    });
  }

  /** In principe "private" members */

  static getAttachment(id, name, size) {
    return new Office.Promise((resolve, reject) => {
      Office.context.mailbox.item.getAttachmentContentAsync(id, async (result) => {
        // eslint-disable-next-line eqeqeq
        if (result.status == Office.AsyncResultStatus.Succeeded) {
          const { format, content } = result.value;
          const headerValues = {
            contentType: undefined,
            contentTransferEncoding: undefined,
          };

          const unknownContentType = "application/octet-stream";
          switch (format) {
            case Office.MailboxEnums.AttachmentContentFormat.Base64: {
              const mime = await fileTypeFromBuffer(Buffer.from(content, "base64"));
              headerValues.contentType = !mime ? unknownContentType : mime.mime;
              headerValues.contentTransferEncoding = "base64";
              break;
            }
            case Office.MailboxEnums.AttachmentContentFormat.Url: {
              // url staat al in de mail
              return resolve(false);
            }
            case Office.MailboxEnums.AttachmentContentFormat.Eml: {
              // falls through
            }
            default: {
              headerValues.contentType = unknownContentType;
              break;
            }
          }

          const headers = [];
          headers.push(`Content-Type: ${headerValues.contentType}; name="${name}"`);
          headers.push(`Content-Disposition: attachment; fileName="${name}"; size=${size};`);
          const hasEncoding = !!headerValues.contentTransferEncoding;
          if (hasEncoding) headers.push(`Content-Transfer-Encoding: ${headerValues.contentTransferEncoding}`);

          resolve({ headers, body: content });
        } else {
          reject(new Error(result.status));
        }
      });
    });
  }

  static getEmailAsHTML(boundary) {
    return new Office.Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync("html", {}, (result) => {
        // eslint-disable-next-line eqeqeq
        if (result.status == Office.AsyncResultStatus.Succeeded) {
          let part = "";
          if (result.value !== "") {
            part += `\r\n\r\n--${boundary}\r\n`;
            part += 'Content-type: text/html; charset="UTF-8"\r\n\r\n';
            part += result.value;
          }
          resolve(part);
        } else {
          reject(new Error(result.status));
        }
      });
    });
  }

  static getEmailAsText(boundary) {
    return new Office.Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync("text", {}, (result) => {
        // eslint-disable-next-line eqeqeq
        if (result.status == Office.AsyncResultStatus.Succeeded) {
          let part = "";
          if (result.value !== "") {
            part += `\r\n\r\n--${boundary}\r\n`;
            part += 'Content-type: text/plain; charset="UTF-8"\r\n\r\n';
            part += result.value;
          }
          resolve(part);
        } else {
          reject(new Error(result.status));
        }
      });
    });
  }
}
