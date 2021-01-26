// office-addin-react - Koppeling van Mozard met Microsoft Office
// Copyright (C) v  Mozard BV
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

import React, { useEffect, useState } from "react";
import { TooltipHost, DefaultButton, TextField } from "office-ui-fabric-react";
import { v4 as uuidv4 } from "uuid";

import Header from "./Header";

function ViewSettings() {
  const [auth, setAuth] = useState("");
  const [authErrorMessage, setAuthErrorMessage] = useState("");
  const [authId] = useState(uuidv4());
  const [env, setEnv] = useState("");
  const [envErrorMessage, setEnvErrorMessage] = useState("");
  const [envId] = useState(uuidv4());
  const [showSuccess, setShowSuccess] = useState("");

  const deleteFnvb = () => {
    localStorage.setItem(
      "currentFnvb",
      JSON.stringify({
        auth: "",
        env: "",
      })
    );
    window.location.reload();
  };
  const saveFnvb = () => {
    if (env.length < 5) {
      setEnvErrorMessage("Fout: ongeldige omgeving");
    }
    if (auth.length < 40) {
      setAuthErrorMessage("Fout: ongeldige officecode");
    }
    if (env.length === 5 && auth.length === 40) {
      setAuthErrorMessage(false);
      setEnvErrorMessage(false);
      setShowSuccess("Je instellingen zijn opgeslagen!");

      localStorage.setItem(
        "currentFnvb",
        JSON.stringify({
          auth,
          env,
        })
      );
    }
  };

  useEffect(() => {
    if (localStorage.getItem("currentFnvb")) {
      const { auth, env } = JSON.parse(localStorage.getItem("currentFnvb"));
      setAuth(auth);
      setEnv(env);
    }
  }, []);

  return (
    <div className="view-settings">
      <Header />
      <form className="mt-4 px-4">
        <TooltipHost
          content="Vul hier de unieke code voor je Mozardomgeving in. Deze kun je opvragen bij je functioneel beheerder."
          id={envId}
        >
          <TextField
            aria-describedby={envId}
            aria-required
            errorMessage={envErrorMessage}
            label="Omgeving"
            maxLength="5"
            onChange={(event) => setEnv(event.target.value)}
            placeholder="Bijv.: mzrdp"
            required
            type="text"
            value={env}
          />
        </TooltipHost>
        <TooltipHost
          content="Vul hier jouw persoonlijke Officecode in. Deze kun je vinden in Mozard, bij je profielopties."
          id={authId}
        >
          <TextField
            aria-describedby={authId}
            aria-required
            errorMessage={authErrorMessage}
            label="Officecode"
            maxLength="40"
            onChange={(event) => setAuth(event.target.value)}
            placeholder="Bijv.: quin1the1yahrieT1phi2Sai0jaicohxiaJieyae"
            required
            type="password"
            value={auth}
          />
        </TooltipHost>
        <DefaultButton className="mt-4 w-100" onClick={() => saveFnvb()} primary text="Opslaan" />
        <DefaultButton className="mt-4 w-100" onClick={() => deleteFnvb()} text="Afmelden" />
      </form>
      {showSuccess && (
        <div className="success text-p-4 center w-100">
          <span aria-hidden="true" className="mr-4 ms-fontSize-24 ms-Icon ms-Icon--Accept"></span>
          <span>{showSuccess}</span>
        </div>
      )}
    </div>
  );
}

export default ViewSettings;
