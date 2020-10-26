import * as React from "react";
import { Spinner, SpinnerSize } from "office-ui-fabric-react";

export default class Progress extends React.Component {
  render() {
    const { message } = this.props;

    return (
      <section className="ms-u-fadeIn500 ms-welcome__progress mt-8">
        <Spinner label={message} type={SpinnerSize.large} />
      </section>
    );
  }
}
