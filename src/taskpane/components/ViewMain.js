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

export default class ViewMain extends React.Component {
  constructor() {
    super();

    this.state = {
      auth: "",
      env: "",
      documentExtension: "",
      documentId: "",
      documentIdFromDocument: undefined,
      documentIdFromDocumentPrevious: undefined,
      documentName: "",
      documentType: "",
      dossierId: "",
      dossierIdFromUser: false,
      initialized: false,
      platform: "",
      progress: {
        description: "",
        label: "Klaar om te verzenden",
        percentComplete: 0,
      },
      responseDocumentTypes: {},
      showError: false,
      showProgress: true,
      showSelectDocumentType: false,
      showSpinner: false,
    };
    this.handleDossierIdChange = this.handleDossierIdChange.bind(this);
    this.handleDocumentIdChange = this.handleDocumentIdChange.bind(this);
    this.handleDocumentNameChange = this.handleDocumentNameChange.bind(this);
    this.handleDocumentTypeChange = this.handleDocumentTypeChange.bind(this);
    this.handlePromptAsNew = this.handlePromptAsNew.bind(this);
    this.sendFile = this.sendFile.bind(this);
    this.submitNew = this.submitNew.bind(this);
  }

  handleDocumentIdChange(event) {
    this.setState({ documentId: event.target.value });
  }

  handleDocumentNameChange(event) {
    this.setState({ documentName: event.target.value });
  }

  handleDocumentTypeChange(event, option) {
    this.setState({ documentType: option.text });
  }

  handleDossierIdChange(event) {
    this.setState({ dossierId: event.target.value });
  }

  handlePromptAsNew() {
    this.setState({
      documentIdFromDocumentPrevious: this.state.documentIdFromDocument,
      documentIdFromDocument: false,
      progress: {
        description: "",
        label: "Klaar om te verzenden",
        percentComplete: 0,
      },
      showError: false,
    });
  }

  preventFormSubmit(event) {
    event.preventDefault();
  }

  sendFile(event) {
    this.setState({
      showError: false,
      showProgress: true,
    });

    const progressCallback = (percentComplete, description, label) => {
      this.setState({
        progress: {
          description,
          label: label || "Bezig met verzenden",
          percentComplete,
        },
        showProgress: true,
      });
    };

    if (!this.state.documentIdFromDocument) {
      OfficeDocument.setDocumentId(this.state.documentId).catch((e) => {
        console.info(e);
      });
    }

    try {
      Middleware.sendFile(
        progressCallback,
        {
          auth: this.state.auth,
          env: this.state.env,
          platform: this.state.platform,
        },
        {
          documentExtension: this.state.documentExtension,
          documentId: this.state.documentId,
          documentName: this.state.documentName,
          documentType: this.state.documentType,
          dossierId: this.state.dossierId,
        }
      );
    } catch (e) {
      console.error(e);
      this.setState({
        progress: {
          description: "",
          percentComplete: undefined,
        },
        showError: e,
        showProgress: false,
      });
    }
  }

  submitNew() {
    this.setState({
      dossierIdFromUser: true,
      showError: false,
      showSpinner: true,
    });

    Middleware.getDocTypes(
      {
        auth: this.state.auth,
        env: this.state.env,
      },
      { dossierId: this.state.dossierId }
    )
      .then((res) => {
        this.setState({
          documentId: res.data.moz_vnr_document,
          responseDocumentTypes: res.data.moz_vnr_documenttypen.moz_vnr_documenttype,
          showSelectDocumentType: true,
          showSpinner: false,
        });
      })
      .catch((e) => {
        this.setState({
          showError: "Fout: Geen privileges",
          showProgress: false,
          showSpinner: false,
        });
        console.error(e);
      });
  }

  componentDidMount() {
    if (localStorage.getItem("currentFnvb")) {
      const currentFnvb = JSON.parse(localStorage.getItem("currentFnvb"));
      this.setState({
        auth: currentFnvb.auth,
        env: currentFnvb.env,
      });
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

    this.setState({ platform: getHostInfo().platform });

    switch (getHostInfo().type) {
      case "Word":
        this.setState({ documentExtension: "docx" });
        break;
      case "Excel":
        this.setState({ documentExtension: "xlsx" });
        break;
      case "Powerpoint":
        this.setState({ documentExtension: "pptx" });
        break;
      default:
        this.setState({ documentExtension: "" });
        break;
    }

    OfficeDocument.getDocumentId()
      .then((res) => {
        this.setState({
          documentId: res.value,
          documentIdFromDocument: true,
          initialized: true,
          progress: {
            description: `Nieuwe versie van d${res.value}`,
            label: "Klaar om te verzenden",
            percentComplete: 0,
          },
        });
      })
      .catch((e) => {
        this.setState({
          documentIdFromDocument: false,
          initialized: true,
        });
      });
  }

  render() {
    return (
      <div className="view-main">
        <Header />
        {this.state.documentIdFromDocument === false && (
          <form className="mt-4 px-4" onSubmit={this.formPreventDefault}>
            <TextField
              label="Zaaknummer (nieuw document)"
              onChange={this.handleDossierIdChange}
              prefix="z"
              type="number"
              value={this.state.dossierId}
            />
            <PrimaryButton
              className="mt-4 w-100"
              iconProps={{ iconName: "Add" }}
              onClick={this.submitNew}
              text="Verzenden als nieuw document"
            />
            {!this.state.documentIdFromDocumentPrevious && !this.state.dossierIdFromUser && (
              <div>
                <hr className="mb-4 mt-8" />
                <TextField
                  label="Documentnummer (nieuwe versie)"
                  onChange={this.handleDocumentIdChange}
                  prefix="d"
                  type="number"
                  value={this.state.documentId}
                />
                <PrimaryButton
                  className="mt-4 w-100"
                  iconProps={{ iconName: "Refresh" }}
                  onClick={this.sendFile}
                  text="Verzenden als nieuwe versie"
                />
              </div>
            )}
          </form>
        )}

        {this.state.documentIdFromDocument && (
          <form className="px-4" onSubmit={this.formPreventDefault}>
            <PrimaryButton
              className="mt-4 w-100"
              iconProps={{ iconName: "Refresh" }}
              onClick={this.sendFile}
              text="Verzenden als nieuwe versie"
            />
            <DefaultButton
              className="mt-4 w-100"
              iconProps={{ iconName: "Add" }}
              onClick={this.handlePromptAsNew}
              text="Verzenden als nieuw document"
            />
          </form>
        )}

        {this.state.showSelectDocumentType && (
          <form className="mt-2 px-4" onSubmit={this.formPreventDefault}>
            <TextField
              defaultValue={this.state.documentName}
              label="Documentnaam"
              onChange={this.handleDocumentNameChange}
              suffix={`.${this.state.documentExtension}`}
              type="text"
            />
            <Dropdown
              label="Documenttype"
              onChange={this.handleDocumentTypeChange}
              options={this.state.responseDocumentTypes.map((type) => {
                return {
                  text: type.moz_doct_naam,
                  key: type.moz_doct_volgnr,
                };
              })}
              responsiveMode="large"
            />
            <PrimaryButton className="mt-4 w-100" onClick={this.sendFile} text="Verzenden" />
          </form>
        )}

        {this.state.showError && (
          <div className="error text-p-4 center w-100">
            <span aria-hidden="true" className="mr-4 ms-fontSize-24 ms-Icon ms-Icon--Error"></span>
            <span>{this.state.showError}</span>
          </div>
        )}

        {this.state.showSpinner && (
          <div className="mt-4 text-center">
            <Spinner size={SpinnerSize.large} />
          </div>
        )}

        {this.state.showProgress && this.state.initialized && (
          <div className="progress text-center w-100">
            <ProgressIndicator
              description={this.state.progress.description}
              label={this.state.progress.label}
              percentComplete={this.state.progress.percentComplete}
            />
          </div>
        )}
      </div>
    );
  }
}
