// office-addin-react - Koppeling van Mozard met Microsoft Office
// Copyright (C) 2021-2022  Mozard BV
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
import "react-app-polyfill/ie11";
import "@fluentui/react";
import "./taskpane.css";
import App from "./components/App";
import { initializeIcons } from "@fluentui/react/lib/Icons";
import * as React from "react";
import * as ReactDOM from "react-dom";
import * as Sentry from "@sentry/react";
import { Integrations } from "@sentry/tracing";

initializeIcons();

let isOfficeInitialized = false;

const title = "Mozard";

Sentry.init({
  dsn: "https://71ac0bb66a8a4bba92810228e8597ad3@o495609.ingest.sentry.io/5578600",
  autoSessionTracking: true,
  integrations: [new Integrations.BrowserTracing()],
  tracesSampleRate: 1.0,
});

const render = (Component) => {
  ReactDOM.render(
    <React.StrictMode>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} />
    </React.StrictMode>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.onReady(() => {
  isOfficeInitialized = true;
  render(App);
});

/* Initial render showing a progress bar */
render(App);

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
