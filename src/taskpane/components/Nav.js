import * as React from "react";
import { CommandBar } from "office-ui-fabric-react";

export default class Nav extends React.Component {
  render() {
    const _items = [
      {
        href: "https://www.mozard.nl",
        iconProps: { iconName: "LightningBolt" },
        key: "fnvb",
        text: "Powered by Mozard",
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
}
