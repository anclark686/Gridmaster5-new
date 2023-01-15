import 'components/check/css/check.css';
import React, { useState, useEffect } from "react";
import { useLocation, Link, useNavigate } from "react-router-dom";
import { post } from 'utils/requests';
import Titlebar from "components/titlebar/Titlebar";
import Navbar from "components/navbar/navbar";
import projectData from "../../json_templates/projectData.json";
import { Spinner, Form, Card, Modal } from "react-bootstrap";

function Check() {
  const countries = projectData.Countries;
  const languages = projectData.Languages;
  const sessionTypes = projectData["Session Type"];
  const abbreviations = projectData.Abbreviations;
  const [projectDetails, setProjectDetails] = useState(null);
  const [loading, setLoading] = useState(false);
  const [fullscreen, setFullscreen] = useState(true);
  const [statusMessage, setStatusMessage] = useState("");
  const [validMessage, setValidMessage] = useState("");

  const [country, setCountry] = useState("");
  const [language, setLanguage] = useState("");
  const [sessType, setSessType] = useState("");
  const [multiRCR, setMultiRCR] = useState(false);
  const [rcrName, setRCRName] = useState("");
  const [projNotes, setProjNotes] = useState([]);

  const [showDate, setShowDate] = useState(false);
  const [pdtCombo, setPDTCombo] = useState([]);
  const [gdtCombo, setGDTCombo] = useState([]);
  const [dateMismatch, setDateMismatch] = useState(false);

  const [showGrid, setShowGrid] = useState(false);
  const [gridIssues, setGridIssues] = useState([]);
  const [showMissing, setShowMissing] = useState(false);

  const navigate = useNavigate();

  let data = useLocation();
  const masterNum = data.state.masterNum;

  const initDropdowns = (details) => {
    setCountry(details.country);
    setLanguage(details.language);
    setSessType(details["session-type"]);
  };

  const getTime = (timeStr) => {
    let [start, end] = timeStr.split(":");
    start = parseInt(start);
    if (end.includes("PM") && start !== 12) {
      start = start + 12;
    }
    end = parseInt(end.slice(0, 5));

    const d = new Date();
    d.setHours(start, end, 0, 0);
    return d;
  };

  const genDateTime = (details) => {
    console.log(details);
    const pdates = details.dates;
    const ptimes = details.times;
    const gdates = details["grid_dates"];
    const gtimes = details["grid_times"];
    let checkedDates = [];
    let checkedTimes = [];
    setPDTCombo(pdates.map((x, i) => [x, ptimes[i]]));
    console.log(pdates.map((x, i) => [x, ptimes[i]]));

    for (let i in gdates) {
      if (pdates.includes(gdates[i])) {
        console.log("or here");
        checkedDates.push("good");
        let location = pdates.indexOf(gdates[i]);

        let [pStart, pEnd] = ptimes[location].split("-");
        let [gStart, gEnd] = gtimes[i].split("-");
        pStart = getTime(pStart);
        pEnd = getTime(pEnd);
        gStart = getTime(gStart);
        gEnd = getTime(gEnd);

        if (gStart.getTime() < pStart.getTime()) {
          checkedTimes.push("outside");
        } else if (gEnd.getTime() > pEnd.getTime()) {
          checkedTimes.push("outside");
        } else {
          checkedTimes.push("good");
        }
      } else {
        checkedDates.push("missing");
        checkedTimes.push("outside");
      }
    }
    setGDTCombo(
      gdates.map((x, i) => [x, gtimes[i], checkedDates[i], checkedTimes[i]])
    );
    console.log(
      gdates.map((x, i) => [x, gtimes[i], checkedDates[i], checkedTimes[i]])
    );

    if (pdates.checkedDates !== 0) {
      if (
        checkedDates.includes("missing") ||
        checkedTimes.includes("outside")
      ) {
        setDateMismatch(true);
      } else {
        setDateMismatch(false);
      }
    }
  };

  const getProjectDetails = async (instructions) => {
    setLoading(true);
    if (instructions === "full") {
      setStatusMessage("Cleaning Grid, and Checking PC...");
    } else if (instructions === "date") {
      setStatusMessage("Checking new dates...");
    } else if (instructions === "grid") {
      setStatusMessage("Rechecking grid...");
    }
    await post(
      [masterNum, instructions], // Body
      'details', // Route
      (res) => {
        setLoading(false);
        console.log(res.data)
        // setStatusMessage("");
        // console.log(res.data);
        // setProjectDetails(res.data);
        // initDropdowns(res.data);
        // genDateTime(res.data);
        // setGridIssues(res.data["grid_notes"]);
        // setProjNotes(res.data["project-notes"].split("\n"));
      }, // Response callback
      (error) => console.error(error) // Error callback
    )
  };

  // useEffect(() => {
  //     getProjectDetails("full")
  // }, [])

  const validateDetails = () => {
    let current_message = validMessage;

    if (!multiRCR) {
      setRCRName("");
      current_message = current_message.includes("Notes")
        ? current_message.replace("RCR not listed in Notes. ", "")
        : current_message;
      current_message = current_message.includes("name")
        ? current_message.replace("No RCR name provided. ", "")
        : current_message;
      setValidMessage(current_message);
    } else {
      if (rcrName === "") {
        current_message = current_message.includes("Notes")
          ? current_message.replace("RCR not listed in Notes. ", "")
          : current_message;
        current_message = !current_message.includes("name")
          ? current_message + "No RCR name provided. "
          : current_message;
        setValidMessage(current_message);
        return false;
      } else {
        setRCRName(rcrName[0].toUpperCase() + rcrName.substring(1));
        console.log(rcrName);
        current_message = current_message.includes("name")
          ? current_message.replace("No RCR name provided. ", "")
          : current_message;
        setValidMessage(current_message);
        if (!projectDetails["project-notes"].includes(rcrName)) {
          current_message = !current_message.includes("Notes")
            ? current_message + "RCR not listed in Notes. "
            : current_message;
          setValidMessage(current_message);
          return false;
        } else {
          current_message = current_message.includes("Notes")
            ? current_message.replace("RCR not listed in Notes. ", "")
            : current_message;
          setValidMessage(current_message);
        }
      }
    }
    if (
      country &&
      language &&
      sessType &&
      gridIssues.length === 0 &&
      !dateMismatch
    ) {
      console.log("you get here??????");
      setValidMessage("good");
      return true;
    } else {
      console.log("you get here?");
      if (!country) {
        current_message = !current_message.includes("Country")
          ? current_message + "Invalid Country. "
          : current_message;
        setValidMessage(current_message);
      } else {
        current_message = current_message.includes("Country")
          ? current_message.replace("Invalid Country. ")
          : current_message;
        setValidMessage(current_message);
      }
      if (!language) {
        current_message = !current_message.includes("Language")
          ? current_message + "Invalid Language. "
          : current_message;
        setValidMessage(current_message);
      } else {
        current_message = current_message.includes("Language")
          ? current_message.replace("Invalid Language. ")
          : current_message;
        setValidMessage(current_message);
      }
      if (!sessType) {
        current_message = !current_message.includes("Session")
          ? current_message + "Invalid Session Type. "
          : current_message;
        setValidMessage(current_message);
      } else {
        current_message = current_message.includes("Session")
          ? current_message.replace("Invalid Session Type. ")
          : current_message;
        setValidMessage(current_message);
      }
      if (gridIssues.length !== 0) {
        current_message = !current_message.includes("Grid")
          ? current_message + "Please correct Grid. "
          : current_message;
        setValidMessage(current_message);
      } else {
        current_message = current_message.includes("Grid")
          ? current_message.replace("Please correct Grid. ")
          : current_message;
        setValidMessage(current_message);
      }
      if (dateMismatch) {
        current_message = !current_message.includes("Dates")
          ? current_message + "Please correct Dates/Times. "
          : current_message;
        setValidMessage(current_message);
      } else {
        current_message = current_message.includes("Dates")
          ? current_message.replace("Please correct Dates/Times. ")
          : current_message;
        setValidMessage(current_message);
      }
      return false;
    }
  };

  const handleDateShow = () => {
    setFullscreen(true);
    setShowDate(true);
  };

  const handleGridShow = () => {
    setFullscreen(true);
    setShowGrid(true);
  };

  const handleMissingShow = () => {
    setShowMissing(true);
  };

  const handleMissingClose = () => {
    setShowMissing(false);
  };

  const goToUpdate = async (instructions) => {
    console.log(validMessage);
    let ready = false;
    if (instructions !== "override") {
      ready = validateDetails();
    } else {
      ready = true;
    }
    console.log(ready);
    const userDetails = {
      country: country,
      language: language,
      session_type: sessType,
      rcr_name: rcrName,
    };
    console.log(userDetails);
    if (ready) {
      await post(
        userDetails, // Body
        'update', // Route
        (res) => {
          navigate("/update", { state: { masterNum: masterNum } });
        }, // Response callback
        (error) => console.error(error) // Error callback
      )
    }
  };

  return (
    <div className="check">
      <Titlebar />
      <Navbar />
      <h1 className="title">{masterNum.toUpperCase()}</h1>
      {projectDetails ? (
        <>
          {projectDetails["grid_notes"].includes("Master mismatch") ? (
            <h5 className="missing">Grid Master Number Mismatch</h5>
          ) : null}
        </>
      ) : null}
      <h4>Project Details</h4>
      {!loading ? (
        <div className="project-details">
          <button onClick={() => getProjectDetails("full")}>
            Get Project Details
          </button>

          <div className="project-notes">
            <h5>Project Notes</h5>
            <Card className="notes-card">
              {projNotes.map((note, index) => (
                <p key={index}>{note}</p>
              ))}
            </Card>
          </div>
          <div className="check-modals">
            <div className="grid-check">
              {gridIssues.length > 0 ? (
                <button
                  className="btn btn-danger btn-lg check-btn"
                  onClick={() => handleGridShow()}
                >
                  Issue with Grid
                </button>
              ) : (
                <button
                  className="btn btn-info btn-lg check-btn"
                  onClick={() => handleGridShow()}
                >
                  Check Update Info
                </button>
              )}
              <Modal
                show={showGrid}
                fullscreen={fullscreen}
                onHide={() => setShowGrid(false)}
              >
                <Modal.Header closeButton>
                  <Modal.Title>Grid Stats and Issues</Modal.Title>
                </Modal.Header>
                <Modal.Body>
                  {projectDetails ? (
                    <>
                      {Object.entries(projectDetails["stats"]).map(
                        ([key, value], index) => (
                          <span className="stats">
                            {key === "missing_rsps" ? (
                              <>
                                <p>
                                  <strong>
                                    Respondents missing from grid:
                                  </strong>{" "}
                                  {value}
                                </p>
                                {value !== 0 ? (
                                  <>
                                    <button
                                      className="btn btn-info"
                                      onClick={() => handleMissingShow()}
                                    >
                                      Check
                                    </button>
                                    {showMissing ? (
                                      <>
                                        <table>
                                          <tr>
                                            <td>Date</td>
                                            <td>Time</td>
                                            <td>Name</td>
                                          </tr>

                                          {projectDetails["missing_rsps"].map(
                                            (rsp, index) => (
                                              <tr>
                                                <td>{rsp["date"]}</td>
                                                <td>{rsp["group-time"]}</td>
                                                <td>{rsp["name"]}</td>
                                              </tr>
                                            )
                                          )}
                                        </table>
                                        <button
                                          className="btn btn-info"
                                          onClick={() => handleMissingClose()}
                                        >
                                          Close
                                        </button>
                                      </>
                                    ) : null}
                                  </>
                                ) : null}
                              </>
                            ) : null}
                            {key === "rsps_to_add" ? (
                              <p>
                                <strong>Respondents to Add:</strong> {value}
                              </p>
                            ) : null}
                            {key === "rsps_to_change" ? (
                              <p>
                                <strong>Respondents to Change:</strong> {value}
                              </p>
                            ) : null}
                            {key === "rsps_to_delete" ? (
                              <p>
                                <strong>Respondents to Delete:</strong> {value}
                              </p>
                            ) : null}
                          </span>
                        )
                      )}

                      {gridIssues.map((issue, index) => (
                        <span className="issue">
                          {issue !== "Master mismatch" ? (
                            <>
                              <Form.Check inline type="checkbox" />
                              <p key={index}>{issue}</p>
                            </>
                          ) : null}
                        </span>
                      ))}

                      {gridIssues.length > 0 ||
                      projectDetails["stats"]["missing_rsps"] > 0 ? (
                        <div className="inside-modal">
                          <button
                            className="btn btn-info btn-lg modal-btn"
                            onClick={() => getProjectDetails("grid")}
                          >
                            Recheck Grid
                          </button>
                        </div>
                      ) : (
                        <div className="inside-modal">
                          <h1>No Grid issues found!</h1>
                          <button
                            className="btn btn-info btn-lg modal-btn"
                            onClick={() => setShowGrid(false)}
                          >
                            Close
                          </button>
                        </div>
                      )}
                    </>
                  ) : null}
                </Modal.Body>
              </Modal>
            </div>

            <div className="date-check">
              {dateMismatch ? (
                <button
                  className="btn btn-danger btn-lg check-btn"
                  onClick={() => handleDateShow()}
                >
                  Date/Time Issues
                </button>
              ) : (
                <button
                  className="btn btn-info btn-lg check-btn"
                  onClick={() => handleDateShow()}
                >
                  Check Dates & Times
                </button>
              )}
              <Modal
                show={showDate}
                fullscreen={fullscreen}
                onHide={() => setShowDate(false)}
              >
                <Modal.Header closeButton>
                  <Modal.Title>Project Dates and Times</Modal.Title>
                </Modal.Header>
                <Modal.Body>
                  <div className="dates">
                    {pdtCombo.length !== 0 ? (
                      <div className="pc-dates">
                        <h5>Dates on FVSS</h5>
                        <table>
                          <thead>
                            <tr>
                              <th>Dates</th>
                              <th>Times</th>
                            </tr>
                          </thead>

                          {pdtCombo.map((dt, index) => (
                            <tbody>
                              <tr key={index}>
                                <td>{dt[0]}</td>
                                <td>{dt[1]}</td>
                              </tr>
                            </tbody>
                          ))}
                        </table>
                      </div>
                    ) : null}

                    {gdtCombo.length !== 0 ? (
                      <div className="pc-dates">
                        <h5>Dates on Worksheet</h5>
                        <table>
                          <thead>
                            <tr>
                              <th>Dates</th>
                              <th>Times</th>
                            </tr>
                          </thead>
                          {gdtCombo.map((dt, index) => (
                            <tbody>
                              <tr key={index}>
                                {dt[2] === "missing" ? (
                                  <>
                                    <td className="missing">{dt[0]}</td>
                                    <td className="missing">{dt[1]}</td>
                                  </>
                                ) : (
                                  <>
                                    <td>{dt[0]}</td>
                                    {dt[3] === "outside" ? (
                                      <td className="missing">{dt[1]}</td>
                                    ) : (
                                      <td>{dt[1]}</td>
                                    )}
                                  </>
                                )}
                              </tr>
                            </tbody>
                          ))}
                        </table>
                      </div>
                    ) : null}
                  </div>
                  {dateMismatch ? (
                    <div className="inside-modal">
                      <button
                        className="btn btn-info btn-lg modal-btn"
                        onClick={() => getProjectDetails("date")}
                      >
                        Recheck Dates
                      </button>
                    </div>
                  ) : (
                    <div className="inside-modal">
                      <h1>Looks Good!</h1>
                      <button
                        className="btn btn-info btn-lg modal-btn"
                        onClick={() => setShowDate(false)}
                      >
                        Close
                      </button>
                    </div>
                  )}
                </Modal.Body>
              </Modal>
            </div>
          </div>
          <div className="project-details-menu">
            <div className="menu-item">
              <label htmlFor="countries">Country</label>
              <Form.Select
                name="countries"
                id="countries"
                onChange={(e) => setCountry(e.target.value)}
              >
                <option value={country} selected>
                  {country}
                </option>
                {countries.map((c, index) => (
                  <>
                    {c !== country ? (
                      <option key={index} value={c}>
                        {" "}
                        {c}{" "}
                      </option>
                    ) : null}
                  </>
                ))}
              </Form.Select>
            </div>
            <div className="menu-item">
              <label htmlFor="languages">Language</label>
              <Form.Select
                name="languages"
                id="languages"
                onChange={(e) => setLanguage(e.target.value)}
              >
                <option value={language} selected>
                  {language}
                </option>
                {languages.map((l, index) => (
                  <>
                    {l !== language ? (
                      <option key={index} value={l}>
                        {" "}
                        {l}{" "}
                      </option>
                    ) : null}
                  </>
                ))}
              </Form.Select>
            </div>
            <div className="menu-item">
              <label htmlFor="sessionTypes">Session Type</label>
              <Form.Select
                name="sessionTypes"
                id="sessionTypes"
                onChange={(e) => setSessType(e.target.value)}
              >
                <option value={sessType} selected>
                  {sessType}
                </option>
                {sessionTypes.map((s, index) => (
                  <>
                    {s !== sessType ? (
                      <option key={index} value={s}>
                        {" "}
                        {s}{" "}
                      </option>
                    ) : null}
                  </>
                ))}
              </Form.Select>
            </div>
          </div>

          <div className="rcr">
            <Form>
              <p className="switch-label">Multi-RCR?</p>
              <Form.Check
                type="switch"
                id="custom-switch"
                onChange={() => setMultiRCR(!multiRCR)}
              />
            </Form>
            {multiRCR ? (
              <>
                <p>Recruiter Name?</p>
                <Form.Control
                  type="text"
                  className="recruiter"
                  onChange={(e) => setRCRName(e.target.value)}
                />
                {rcrName}
              </>
            ) : null}
          </div>

          <>
            <h5>{validMessage}</h5>
            {validMessage.length === 25 ? (
              <button
                className="btn btn-info btn-lg"
                onClick={() => goToUpdate("override")}
              >
                Continue anyways?
              </button>
            ) : (
              <button
                className="btn btn-info btn-lg"
                onClick={() => goToUpdate("first")}
              >
                Continue
              </button>
            )}
          </>
        </div>
      ) : (
        <>
          <h4>{statusMessage}</h4>
          <Spinner animation="border" variant="info" />
        </>
      )}
    </div>
  );
}

export default Check;
