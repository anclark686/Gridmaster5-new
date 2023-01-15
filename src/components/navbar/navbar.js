import 'components/navbar/css/Navbar.css';
import React, { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { get } from 'utils/requests';
import { Button, Modal } from 'react-bootstrap';



const logo = require('./img/Forsta.jpg');

function Navbar() {
  const [show, setShow] = useState(false);
  const handleClose = () => setShow(false);
  const handleShow = () => setShow(true);
  const navigate = useNavigate();

  const startOver = async () => {
    const location = window.location.pathname;
    if (location === '/') {
      window.location.reload(false);
    } else {
      navigate('/');
    }
    await get(
      'startover', // Route
      null, // Response callback
      (error) => console.error(error) // Error callback
    );
  };

  return (
    <div className="header">
      <div className="logo">
        <img src={logo} alt="forsta logo" />
      </div>
      <div className="header-buttons">
        <div className="start-over">
          <button className="btn btn-info" onClick={startOver}>
            Start Over
          </button>
        </div>
        <Button variant="info" onClick={handleShow}>
          Help
        </Button>
      </div>
      <Modal size="lg" show={show} onHide={handleClose}>
        <Modal.Header closeButton>
          <Modal.Title>
            <strong>Helpful Tips & Tricks</strong>
          </Modal.Title>
        </Modal.Header>
        <Modal.Body
          style={{
            maxHeight: "500px",
            overflowY: "auto",
          }}
        >
          <h5>
            <strong>Getting Started</strong>
          </h5>
          <p>
            Upon starting the app, a new browser window will open. Use your FV
            Database credentials to login. A folder called "GridMaster" will
            appear on your desktop if this is your first time. To begin, save
            the needed grid with the master number - be sure the file extension
            is .xlsx (ex: M123456.xlsx).
          </p>
          <br />
          <h5>
            <strong>Using the Program</strong>
          </h5>
          <ul>
            <li>
              Once the file is saved, enter the Master Number and hit Start. The
              program will look for the project in FVSS, and will confirm that a
              file exists.
            </li>
            <li>
              Once confirmed, hit Continue. The program will then pull project
              details and respondent data from FVSS & FVPC. It will also parse
              the grid for respondent update information.
            </li>
            <li>
              This next screen will be a page for you to verify the project
              details, and adjust any dates and/or times in FVSS if need be.
            </li>
            <li>
              After the necessary details and dates have been added and
              adjusted, the program will offer a summary of changes. You can
              then review these changes before applying any updates. Once
              approved, hit "Master That Grid". A "Cleaned" version of the grid
              will also open for review.
            </li>
            <li>
              The program will then make any updates to PC, and will return a
              summary of changes made. From here, hit "Continue to Email" to
              email the respondents.
            </li>
            <li>
              Emails will then be sent to respondents highlighted in yellow. A
              summary of emails sent will appear once complete.
            </li>
            <li>
              After emails have been sent, there are no further actions. To
              process another master, hit "Start Over"
            </li>
          </ul>
          <br />
          <h5>
            <strong>Issues</strong>
          </h5>
          <br />
          <h5>
            <strong>GDPR Note</strong>
          </h5>
          <p>
            Any files older than 90 days will be deleted from this folder upon
            running the app.
          </p>
        </Modal.Body>
        <Modal.Footer>
          <Button variant="info" onClick={handleClose}>
            OK
          </Button>
        </Modal.Footer>
      </Modal>
    </div>
  );
}

export default Navbar;
