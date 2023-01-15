import React from "react"
import { Router, Route } from 'electron-router-dom';

import Main from 'components/main/main';
import Check from 'components/check/check';
import Update from 'components/update/update';
import Email from 'components/email/email';

export function AppRoutes() {
  return (
    <Router
      main={ <Route path="/" element={ <Main /> } /> }
      check={ <Route path="check" element={ <Check /> } /> }
      update={ <Route path="update" element={ <Update /> } /> }
      email={ <Route path="email" element={ <Email /> } /> }
    />
  )
}