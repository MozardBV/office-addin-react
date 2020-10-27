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

import * as React from "react";
import { HashRouter as Router, Route } from "react-router-dom";

import Header from "./Header";
import Progress from "./Progress";
import Nav from "./Nav";
import ViewMain from "./ViewMain";
import ViewSettings from "./ViewSettings";

export default class App extends React.Component {
  render() {
    const { isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <div>
          <Header />
          <Progress message="Deze add-in werkt uitsluitend in Microsoft Office." />
        </div>
      );
    }

    if (isOfficeInitialized) {
      return (
        <div id="app">
          <Router>
            <Route exact path="/" component={ViewMain} />
            <Route exact path="/settings" component={ViewSettings} />
          </Router>
          <Nav />
        </div>
      );
    }
  }
}
