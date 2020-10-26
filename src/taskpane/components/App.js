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
