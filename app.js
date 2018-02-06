/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
// application dependencies
const express = require('express');
const session = require('express-session');
const path = require('path');
const favicon = require('serve-favicon');
const logger = require('morgan');
const cookieParser = require('cookie-parser');
const bodyParser = require('body-parser');
const routes = require('./routes/index');
const passport = require('passport');
const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
const uuid = require('uuid');
const config = require('./utils/config.js');

const app = express();

// **IMPORTANT
// Note that production apps will need to create a self-signed cert and use a secure server,
// and change dev settings marked 'For development only' in app.js and config.js.
// Below is an example after you have the key cert pair:

// const https = require('https');
// const fs = require('fs');
// const certConfig = {
//  key: fs.readFileSync('./key.pem', 'utf8'),
//  cert: fs.readFileSync('./cert.pem', 'utf8')
// };
// const server = https.createServer(certConfig, app);

// authentication setup
const callback = (iss, sub, profile, accessToken, refreshToken, done) => {
  console.log(accessToken)
  done(null, {
    profile,
    accessToken,
    refreshToken
  });
};

passport.use(new OIDCStrategy(config.creds, callback));

const users = {};
passport.serializeUser((user, done) => {
  const id = uuid.v4();
  users[id] = user;
  done(null, id);
});
passport.deserializeUser((id, done) => {
  const user = users[id];
  done(null, user);
});

process.env.NODE_ENV= "development"

// view engine setup
app.set('views', path.join(__dirname, 'views'));

// set the view engine to ejs
app.set('view engine', 'ejs');

app.use(favicon(path.join(__dirname, 'public', 'img', 'favicon.ico')));
app.use(logger('dev'));
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));
app.use(cookieParser());

// session middleware configuration
// see https://github.com/expressjs/session
app.use(session({
  secret: '12345QWERTY-SECRET',
  name: 'graphNodeCookie',
  resave: false,
  saveUninitialized: false,
  //cookie: {secure: true} // For development only
}));

app.use(express.static(path.join(__dirname, 'public')));
app.use(passport.initialize());
app.use(passport.session());
app.use('/', routes);

// error handlers
// catch 404 and forward to error handler
app.use(function (req, res, next) {
  const err = new Error('Not Found');
  //console.log(err);
  res.status(404);
  next(err);
});

// development error handler
// will print stacktrace
if (app.get('env') === 'development') {
  app.use(function (err, req, res) {
    res.status(err.status || 500);
    res.render('error', {
      message: err.message,
      error: err
    });
  });
}

app.get('/login', passport.authenticate('azuread-openidconnect',{ failureRedirect:'/'}),
(req, res, next) => {
//  const accessToken = req.user.accessToken; // user = req.user.profile.displayName;
   console.log("Access token is generated in login route: "+ req.user);
   console.log(app.get('accessToken'));
});


// production error handler
// no stacktraces leaked to user
app.use(function (err, req, res) {
  res.status(err.status || 500);
  res.render('error', {
    message: err.message,
    error: {}
  });
});

module.exports = app;
