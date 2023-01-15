import 'components/start/css/start.css';
import React, { useState, useEffect } from "react";
import { Link, useNavigate } from "react-router-dom";
import { Spinner, Card, Form, FloatingLabel } from "react-bootstrap";
import { post } from 'utils/requests';

function Start() {
  const [masterNum, setMasterNum] = useState("");
  const [validMaster, setValidMaster] = useState(false);
  const [found, setFound] = useState("");
  const [masterPopulated, setMasterPopulated] = useState(false);
  const [hitStart, setHitStart] = useState(false);
  const [loading, setLoading] = useState(false);
  const navigate = useNavigate();
  const mregex = /[m][1-9]{6}$/i;
  const nregex = /[1-9]{6}$/;

  const validateMaster = () => {
    if (masterNum.slice(0, 1) === "m") {
      setValidMaster(mregex.test(masterNum) ? true : false);
    } else {
      setValidMaster(nregex.test(masterNum) ? true : false);
      setMasterNum("m" + masterNum);
    }
  };

  useEffect(() => {
    setHitStart(false);
    if (masterNum.length >= 6) {
      setMasterPopulated(true);
      validateMaster();
    } else {
      setMasterPopulated(false);
      setValidMaster(false);
      setFound('');
    }
  }, [masterNum]);

  const checkMaster = async () => {
    setLoading(true);
    setHitStart(true);
    validateMaster(masterNum);
    if (validMaster) {
      await post(
        masterNum, // Body
        'master', // Route
        (response) => {
          setLoading(false);
          console.log(response.data);
          setFound(response.data);
          console.log(found);
        }, // Response callback
        (error) => console.error(error) // Error callback
      )
    } else {
      setLoading(false);
    }
  };

  const cleanAndCheck = async () => {
    validateMaster(masterNum);
    try {
      navigate("check", { state: { masterNum: masterNum } });
    } catch (err) {
      console.log(err);
    }
  };

  const handleKeyDown = (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      checkMaster();
    }
  };

  return (
    <div className="start">
      <FloatingLabel
        controlId="floatingInput"
        label="Master Number"
        className="master-field"
      >
        <Form.Control
          type="text"
          placeholder="MXXXXXX"
          onChange={(e) => setMasterNum(e.target.value.toLowerCase())}
          onKeyDown={handleKeyDown}
        />
      </FloatingLabel>
      {!loading ? (
        <>
          {!found ? (
            <button
              className="btn btn-info btn-lg start-button"
              onClick={() => checkMaster()}
            >
              Start
            </button>
          ) : null}
        </>
      ) : (
        <Spinner animation="border" variant="info" />
      )}
      {!masterPopulated && hitStart ? (
        <h5>Please enter a valid Master Number</h5>
      ) : (
        <>
          {masterPopulated && hitStart ? (
            <>
              {found === "File and Project found" && validMaster ? (
                <div className="found-continue">
                  <h5>File and Project found</h5>
                  <button
                    className="btn btn-info btn-lg"
                    onClick={() => cleanAndCheck()}
                  >
                    Continue
                  </button>
                </div>
              ) : (
                <div className="invalid">
                  {validMaster ? (
                    <div className="not-found">
                      <h5>{found}</h5>
                    </div>
                  ) : (
                    <div className="not-found">
                      <h5>Invalid Master Number</h5>
                    </div>
                  )}
                </div>
              )}
            </>
          ) : null}
        </>
      )}
    </div>
  );
}

export default Start;
