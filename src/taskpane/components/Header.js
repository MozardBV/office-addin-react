import * as React from "react";
import Logo from "./Logo";

export default class Header extends React.Component {
  render() {
    return (
      <header className="ms-bgColor-neutralLighter ms-u-fadeIn500 ms-welcome__header">
        <Logo />
        <h1 className="ms-fontColor-neutralPrimary ms-fontSize-su ms-fontWeight-light">Officeintegratie</h1>
      </header>
    );
  }
}
