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
import { CommandBar } from "@fluentui/react";

function Nav() {
  const _items = [
    {
      href: "https://www.mozard.nl/mozard/!suite86.scherm0325?mVrg=7781",
      iconProps: { iconName: "Link12" },
      key: "fnvb",
      text: "Meer over Mozard",
    },
  ];

  const _farItems = [
    {
      href: "/#/",
      iconProps: { iconName: "Home" },
      key: "home",
    },
    {
      href: "/#/settings",
      iconProps: { iconName: "Settings" },
      key: "instellingen",
    },
  ];

  return (
    <CommandBar
      ariaLabel="Gebruik de pijltjestoetsen links en rechts om te navigeren"
      className="nav"
      farItems={_farItems}
      items={_items}
    />
  );
}

export default Nav;
