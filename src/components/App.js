import React, { Fragment } from 'react';
import { Routes, Route } from "react-router-dom";


import Main from 'components/main/main';
import Check from 'components/check/check';
import Update from 'components/update/update';
import Email from 'components/email/email';

import styles from 'components/App.module.scss';

function App() {

  return (
    <Fragment>
      <div className={ styles.app }>
        <Routes>
          <Route path="/" element={<Main />} />
          <Route path="update" element={<Update />} />
          <Route path="email" element={<Email />} />
          <Route path="check" element={<Check />} />
        </Routes>
      </div>
    </Fragment>
  );
}

export default App;
