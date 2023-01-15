import React, { useState, useEffect } from 'react';
import { Link } from 'react-router-dom'
import Titlebar from "components/titlebar/Titlebar";
import Navbar from "components/navbar/navbar";

function Update() {

    

    return(
        <div className='update'>
            <Titlebar />
      <Navbar />

            <h1>Update</h1>
            
        </div>
    )
}

export default Update;