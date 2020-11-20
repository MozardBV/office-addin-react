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

import React, { useEffect, useState } from "react";
import {
  DefaultButton,
  Dropdown,
  PrimaryButton,
  ProgressIndicator,
  Spinner,
  SpinnerSize,
  TextField,
} from "office-ui-fabric-react";

import Header from "./Header";

import Middleware from "../api/Middleware";
import OfficeDocument from "../api/OfficeDocument";

function ViewMain() {
  const [auth, setAuth] = useState("");
  const [env, setEnv] = useState("");
  const [documentExtension, setDocumentExtension] = useState("");
  const [documentId, setDocumentId] = useState("");
  const [documentIdErrorMessage, setDocumentIdErrorMessage] = useState("");
  const [documentIdFromDocument, setDocumentIdFromDocument] = useState(undefined);
  const [documentIdFromDocumentPrevious, setDocumentIdFromDocumentPrevious] = useState(undefined);
  const [documentName, setDocumentName] = useState("");
  const [documentNameErrorMessage, setDocumentNameErrorMessage] = useState("");
  const [documentType, setDocumentType] = useState("");
  const [documentTypeErrorMessage, setDocumentTypeErrorMessage] = useState("");
  const [dossierId, setDossierId] = useState("");
  const [dossierIdErrorMessage, setDossierIdErrorMessage] = useState("");
  const [dossierIdFromUser, setDossierIdFromUser] = useState("");
  const [initialized, setInitialized] = useState(false);
  const [platform, setPlatform] = useState("");
  const [progress, setProgress] = useState({
    description: "",
    label: "Klaar om te verzenden",
    percentComplete: 0,
  });
  const [responseDocumentTypes, setResponseDocumentTypes] = useState({});
  const [showError, setShowError] = useState(false);
  const [showProgress, setShowProgress] = useState(false);
  const [showSelectDocumentType, setShowSelectDocumentType] = useState(false);
  const [showSpinner, setShowSpinner] = useState(false);

  const handlePromptAsNew = () => {
    if (!auth || !env) {
      setShowError(
        "Fout: Geen functieverband en/of omgeving gekoppeld. Koppel een functieverband bij het tandwiel rechtsonder."
      );
      setShowProgress(false);
      return;
    }

    setDocumentIdFromDocumentPrevious(documentIdFromDocument);
    setDocumentIdFromDocument(false);
    setProgress({
      description: "",
      label: "Klaar om te verzenden",
      percentComplete: 0,
    });
    setShowError(false);
  };

  const sendFile = (event) => {
    setDocumentIdErrorMessage("");
    setDocumentNameErrorMessage("");
    setDocumentTypeErrorMessage("");

    if (!auth || !env) {
      setShowError(
        "Fout: Geen functieverband en/of omgeving gekoppeld. Koppel een functieverband bij het tandwiel rechtsonder."
      );
      setShowProgress(false);
      return;
    }

    if (!documentId) {
      setDocumentIdErrorMessage("Fout: geen of ongeldig documentnummer opgegeven");
      return;
    }

    if (Object.keys(responseDocumentTypes).length !== 0 && !documentName) {
      setDocumentNameErrorMessage("Fout: geen documentnaam opgegeven");
      return;
    }

    if (documentName) {
      const disallowedChars = ["\\", "*", '"', "<", ">", "|", "%"];
      const invalid = [];

      disallowedChars.forEach((char) => {
        if (documentName.includes(char)) {
          invalid.push(char);
        }
      });

      if (invalid.length > 0) {
        console.log(invalid);
        setDocumentNameErrorMessage(`Fout: ongeldige tekens in documentnaam (${invalid.join(", ")})`);
        return;
      }
    }

    if (Object.keys(responseDocumentTypes).length !== 0 && !documentType) {
      setDocumentTypeErrorMessage("Fout: geen documenttype opgegeven");
      return;
    }

    setShowError(false);
    setShowProgress(true);

    const progressCallback = (percentComplete, description, label) => {
      setProgress({
        description,
        label: label || "Bezig met verzenden",
        percentComplete,
      });
      setShowProgress(true);
    };

    if (!documentIdFromDocument) {
      OfficeDocument.setDocumentId(documentId).catch((e) => {
        console.info(e);
      });
    }

    try {
      Middleware.sendFile(
        progressCallback,
        {
          auth,
          env,
          platform,
        },
        {
          documentExtension,
          documentId,
          documentName,
          documentType,
          dossierId,
        }
      );
    } catch (e) {
      console.error(e);
      setProgress({
        description: "",
        percentComplete: undefined,
      });
      setShowError(e);
      setShowProgress(false);
    }
  };

  const submitNew = () => {
    setDossierIdErrorMessage("");

    if (!dossierId) {
      setDossierIdErrorMessage("Fout: geen of ongeldig zaaknummer opgegeven");
      return;
    }

    setDossierIdFromUser(true);
    setShowError(false);
    setShowSpinner(true);

    Middleware.getDocTypes(
      {
        auth,
        env,
      },
      { dossierId }
    )
      .then((res) => {
        setDocumentId(res.data.moz_vnr_document);
        setResponseDocumentTypes(res.data.moz_vnr_documenttypen.moz_vnr_documenttype);
        setShowSelectDocumentType(true);
        setShowSpinner(false);
      })
      .catch((e) => {
        setShowError("Fout: Geen privileges");
        setShowProgress(false);
        setShowSpinner(false);
        console.error(e);
      });
  };

  useEffect(() => {
    if (localStorage.getItem("currentFnvb")) {
      const { auth, env } = JSON.parse(localStorage.getItem("currentFnvb"));
      setAuth(auth);
      setEnv(env);
    }

    const getHostInfo = () => {
      const hostInfoValue = sessionStorage.getItem("hostInfoValue") || "";

      let items = hostInfoValue.split("$");
      if (items.length < 3) {
        items = hostInfoValue.split("|");
      }

      return {
        type: items[0],
        platform: items[1],
        version: items[2],
      };
    };

    setPlatform(getHostInfo().platform);

    switch (getHostInfo().type) {
      case "Word":
        setDocumentExtension("docx");
        break;
      case "Excel":
        setDocumentExtension("xlsx");
        break;
      case "Powerpoint":
        setDocumentExtension("pptx");
        break;
      default:
        setDocumentExtension("");
        break;
    }

    OfficeDocument.getDocumentId()
      .then((res) => {
        setDocumentId(res.value);
        setDocumentIdFromDocument(true);
        setInitialized(true);
        setProgress({
          description: `Nieuwe versie van d${res.value}`,
          label: "Klaar om te verzenden",
          percentComplete: 0,
        });
      })
      .catch((e) => {
        setDocumentIdFromDocument(false);
        setInitialized(true);
      });
  }, []);

  return (
    <div className="view-main">
      <Header />
      {documentIdFromDocument === false && (
        <form className="mt-4 px-4" onSubmit={(event) => event.preventDefault()}>
          <TextField
            aria-required
            errorMessage={dossierIdErrorMessage}
            label="Zaaknummer (nieuw document)"
            onChange={(event) => setDossierId(event.target.value)}
            prefix="z"
            required
            type="number"
            value={dossierId}
          />
          <PrimaryButton
            className="mt-4 w-100"
            iconProps={{ iconName: "Add" }}
            onClick={() => submitNew()}
            text="Verzenden als nieuw document"
          />
          {!documentIdFromDocumentPrevious && !dossierIdFromUser && (
            <div>
              <hr className="mb-4 mt-8" />
              <TextField
                aria-required
                errorMessage={documentIdErrorMessage}
                label="Documentnummer (nieuwe versie)"
                onChange={(event) => setDocumentId(event.target.value)}
                prefix="d"
                required
                type="number"
                value={documentId}
              />
              <PrimaryButton
                className="mt-4 w-100"
                iconProps={{ iconName: "Refresh" }}
                onClick={() => sendFile()}
                text="Verzenden als nieuwe versie"
              />
            </div>
          )}
        </form>
      )}

      {documentIdFromDocument && (
        <form className="px-4" onSubmit={(event) => event.preventDefault()}>
          <PrimaryButton
            className="mt-4 w-100"
            iconProps={{ iconName: "Refresh" }}
            onClick={(event) => sendFile(event)}
            text="Verzenden als nieuwe versie"
          />
          <DefaultButton
            className="mt-4 w-100"
            iconProps={{ iconName: "Add" }}
            onClick={(event) => handlePromptAsNew(event)}
            text="Verzenden als nieuw document"
          />
        </form>
      )}

      {showSelectDocumentType && (
        <form className="mt-2 px-4" onSubmit={(event) => event.preventDefault()}>
          <TextField
            aria-required
            defaultValue={documentName}
            errorMessage={documentNameErrorMessage}
            label="Documentnaam"
            onChange={(event) => setDocumentName(event.target.value)}
            required
            suffix={`.${documentExtension}`}
            type="text"
          />
          <Dropdown
            errorMessage={documentTypeErrorMessage}
            label="Documenttype"
            onChange={(event, option) => setDocumentType(option.text)}
            options={responseDocumentTypes.map((type) => {
              return {
                text: type.moz_doct_naam,
                key: type.moz_doct_volgnr,
              };
            })}
            responsiveMode="large"
          />
          <PrimaryButton className="mt-4 w-100" onClick={() => sendFile()} text="Verzenden" />
        </form>
      )}

      {progress.percentComplete === 100 && Object.keys(responseDocumentTypes).length !== 0 && (
        <div className="px-4">
          <DefaultButton className="mt-4 w-100" onClick={() => window.location.reload()} text="Terug naar start" />
        </div>
      )}

      {showError && (
        <div className="error text-p-4 center w-100">
          <span aria-hidden="true" className="mr-4 ms-fontSize-24 ms-Icon ms-Icon--Error"></span>
          <span>{showError}</span>
        </div>
      )}

      {showSpinner && (
        <div className="mt-4 text-center">
          <Spinner size={SpinnerSize.large} />
        </div>
      )}

      {showProgress && initialized && (
        <div className="progress text-center w-100">
          <ProgressIndicator
            description={progress.description}
            label={progress.label}
            percentComplete={progress.percentComplete}
          />
        </div>
      )}
    </div>
  );
}

export default ViewMain;
