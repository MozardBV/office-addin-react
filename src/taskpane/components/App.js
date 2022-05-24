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

import React from "react";
import { HashRouter as Router, Routes, Route } from "react-router-dom";
import { Spinner, SpinnerSize } from "@fluentui/react";

import Header from "./Header";
import Nav from "./Nav";
import ViewMain from "./ViewMain";
import ViewSettings from "./ViewSettings";

function App(props) {
  const { isOfficeInitialized } = props;

  if (!isOfficeInitialized) {
    return (
      <div>
        <Header />
        <section className="ms-u-fadeIn500 ms-welcome__progress mt-8">
          <Spinner label="Deze add-in werkt uitsluitend in Microsoft Office." type={SpinnerSize.large} />
        </section>
      </div>
    );
  }

  if (isOfficeInitialized) {
    return (
      <div id="app">
        <Router>
          <Routes>
            <Route exact path="/" element={<ViewMain />} />
            <Route exact path="/settings" element={<ViewSettings />} />
          </Routes>
        </Router>
        <Nav />
      </div>
    );
  }
}

export default App;
