import React from "react"
import Titlebar from "components/titlebar/Titlebar";
import Navbar from "components/navbar/navbar";
import Start from "components/start/start";

function Main() {
  return (
    <div className="main">
      <Titlebar />
      <Navbar />
      <h1 className="title">Gridmaster 5.0</h1>
      <Start />
    </div>
  );
}

export default Main;
