import React, { useState, useEffect } from "react";
import { Link } from "react-router-dom";
import Titlebar from "components/titlebar/Titlebar";
import Navbar from "components/navbar/navbar";

function Email() {
  return (
    <div className="email">
      <Titlebar />
      <Navbar />
      <Link to={"/"}>
        <button className="btn btn-info">Back</button>
      </Link>
      <h1>Email</h1>
    </div>
  );
}

export default Email;
