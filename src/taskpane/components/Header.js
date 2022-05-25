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
import Logo from "./Logo";

function Header() {
  return (
    <header className="ms-bgColor-neutralLighter ms-u-fadeIn500 ms-welcome__header">
      <Logo />
      <h1 className="ms-fontColor-neutralPrimary ms-fontSize-su ms-fontWeight-light">Officeintegratie</h1>
    </header>
  );
}

export default Header;
