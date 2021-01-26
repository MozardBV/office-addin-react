// office-addin-react - Koppeling van Mozard met Microsoft Office
// Copyright (C) 2020  Mozard BV
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

export default class OutlookMailbox {
  static getSubject() {
    return Office.context.mailbox.item.subject;
  }

  static getEmail() {
    return new Office.Promise((resolve, reject) => {
      const boundary = Date.now().toString();

      Office.context.mailbox.item.getAllInternetHeadersAsync((headers) => {
        // Content-Type eraf halen, en er zelf eentje zetten.
        const arr = headers.value.split("\n");
        const indexContentType = arr.findIndex((el) => el.toLowerCase().startsWith("content-type"));
        const indexNextHeader = arr.findIndex((el, idx) => idx > indexContentType && el.includes(":"));

        const arr1 = arr.slice(0, indexContentType);
        const arr2 = arr.slice(indexNextHeader);

        headers = arr1.concat(arr2).join("\n");

        headers += `Content-type: multipart/alternative; boundary="${boundary}"\r\n`;
        headers += "\r\n";

        Promise.all([this.getEmailAsText(boundary), this.getEmailAsHTML(boundary)])
          .then((res) => {
            let body = res.join("");
            body += `\r\n--${boundary}--`;
            resolve({
              counter: 0,
              file: headers.concat(body),
              sliceCount: 1,
            });
          })
          .catch((e) => {
            reject(e);
          });
      });
    });
  }

  /** In principe "private" members */

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
