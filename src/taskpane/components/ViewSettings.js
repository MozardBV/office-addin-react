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
import { TooltipHost, DefaultButton, TextField } from "office-ui-fabric-react";
import { v4 as uuidv4 } from "uuid";

import Header from "./Header";

export default class ViewSettings extends React.Component {
  constructor() {
    super();

    this.state = {
      auth: "",
      authId: uuidv4(),
      env: "",
      envId: uuidv4(),
      authErrorMessage: "",
      envErrorMessage: "",
    };
    this.handleAuthChange = this.handleAuthChange.bind(this);
    this.handleEnvChange = this.handleEnvChange.bind(this);
    this.saveFnvb = this.saveFnvb.bind(this);
  }

  handleAuthChange(event) {
    this.setState({ auth: event.target.value });
  }

  handleEnvChange(event) {
    this.setState({ env: event.target.value });
  }

  saveFnvb() {
    if (this.state.env.length < 5) {
      this.setState({ envErrorMessage: "Fout: ongeldige omgeving" });
    }
    if (this.state.auth.length < 40) {
      this.setState({ authErrorMessage: "Fout: ongeldige officecode" });
    }
    if (this.state.env.length === 5 && this.state.auth.length === 40) {
      this.setState({
        authErrorMessage: false,
        envErrorMessage: false,
      });
      localStorage.setItem(
        "currentFnvb",
        JSON.stringify({
          auth: this.state.auth,
          env: this.state.env,
        })
      );
    }
  }

  componentDidMount() {
    if (localStorage.getItem("currentFnvb")) {
      const currentFnvb = JSON.parse(localStorage.getItem("currentFnvb"));
      this.setState({
        auth: currentFnvb.auth,
        env: currentFnvb.env,
      });
    }
  }

  render() {
    return (
      <div className="view-settings">
        <Header />
        <form className="mt-4 px-4" onSubmit={this.formPreventDefault}>
          <TooltipHost
            content="Vul hier de unieke code voor je Mozardomgeving in. Deze kun je opvragen bij je functioneel beheerder."
            id={this.state.envId}
          >
            <TextField
              aria-describedby={this.state.envId}
              aria-required
              errorMessage={this.state.envErrorMessage}
              label="Omgeving"
              maxLength="5"
              onChange={this.handleEnvChange}
              placeholder="Bijv.: mzrdp"
              required
              type="text"
              value={this.state.env}
            />
          </TooltipHost>
          <TooltipHost
            content="Vul hier jouw persoonlijke Officecode in. Deze kun je vinden in Mozard, bij je profielopties."
            id={this.state.authId}
          >
            <TextField
              aria-describedby={this.state.authId}
              aria-required
              errorMessage={this.state.authErrorMessage}
              label="Officecode"
              maxLength="40"
              onChange={this.handleAuthChange}
              placeholder="Bijv.: quin1the1yahrieT1phi2Sai0jaicohxiaJieyae"
              required
              type="password"
              value={this.state.auth}
            />
          </TooltipHost>
          <DefaultButton className="mt-4 w-100" onClick={this.saveFnvb} primary text="Opslaan" />
        </form>
      </div>
    );
  }
}
