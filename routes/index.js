/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
* This sample shows how to:
*    - Get the current user's metadata
*    - Get the current user's profile photo
*    - Attach the photo as a file attachment to an email message
*    - Upload the photo to the user's root drive
*    - Get a sharing link for the file and add it to the message
*    - Send the email
*/
const express = require('express');
const router = express.Router();
const graphHelper = require('../utils/graphHelper.js');
const emailer = require('../utils/emailer.js');
const passport = require('passport');

const open = require ('open')
// ////const fs = require('fs');
// ////const path = require('path');

// Get the home page.
router.get('/', (req, res) => {
  // check if user is authenticated
  if (!req.isAuthenticated()) {
    res.render('login');
  } else {
    res.redirect('/disconnect')
  }
});

// Authentication request.
router.get('/login',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }),
    (req, res) => {
      res.redirect('/');
    });

// Authentication callback.
// After we have an access token, get user data and load the sendMail page.
router.get('/token',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }),
    (req, res) => {
      graphHelper.getUserData(req.user.accessToken, (err, user) => {
        if (!err) {
          //  req.user.profile.displayName = user.body.displayName;
          //  req.user.profile.emails = [{ address: user.body.mail || user.body.userPrincipalName }];
          // renderSendMail(req, res);
          res.redirect('https://172.16.21.215:3001/view/index.html?access_token='+ req.user.accessToken);
          //shareLink (req, res)
          // graphHelper.getDriveFileList(req.user.accessToken, (err, list) => {
          //   console.log(list)
          // })
        } else {
          renderError(err, res);
        }
    });
  }
)

router.post('/question/launch', (req, res) => {
  console.log(req.body);
  var accessToken = req.body.accessToken, questionId=req.body.questionId, mode=req.body.mode;

  var data = {
     "values": [["id", questionId]]
   }
   console.log(data);
   
  graphHelper.getDriveFileList(accessToken, (err, file) => {
    if (err) { console.log({ err: err.message }) }
   //console.log(JSON.stringify(file));
    console.log('Copying file With ID: ' + file.value[0].id)

    graphHelper.copyFileFromDrive(accessToken, file.value[0].id, (err, id) => {
      console.log('Copied new file ID: ' + id)

      graphHelper.insertDataToExcel(accessToken, id, data, (err, data) => {
        if (err) { console.log({ err: err.message }) }

        console.log('Response after insertion in sheet: \n' + data)

        graphHelper.getSharingLink(accessToken, id, (err, link) => {
           
          if (err) { console.log({ err: err.message }) }
         // open(link)
          console.log('Copied file is available on URL: \n' + link)

          res.json({shareUrl: link});
          // res.render('fileList', {
          //   user: user,
          //   link: link
          // })

        })
      })      
    })
  })
});


router.get('/disconnect', (req, res) => {
  req.session.destroy(() => {
    req.logOut();
    res.clearCookie('graphNodeCookie');
    res.status(200);
    res.redirect('/');
  });
});

// helpers
function hasAccessTokenExpired(e) {
  let expired;
  if (!e.innerError) {
    expired = false;
  } else {
    expired = e.forbidden &&
      e.message === 'InvalidAuthenticationToken' &&
      e.response.error.message === 'Access token has expired.';
  }
  return expired;
}
/**
 * 
 * @param {*} e 
 * @param {*} res 
 */
function renderError(e, res) {
  e.innerError = (e.response) ? e.response.text : '';
  res.render('error', {
    error: e
  });
}

module.exports = router;






// Load the sendMail page.
// function renderSendMail(req, res) {
//   res.render('sendMail', {
//     display_name: req.user.profile.displayName,
//     email_address: req.user.profile.emails[0].address
//   });
// }

// Do prep before building the email message.
// The message contains a file attachment and embeds a sharing link to the file in the message body.
/*
function prepForEmailMessage(req, callback) {
  const accessToken = req.user.accessToken;
  const displayName = req.user.profile.displayName;
  const destinationEmailAddress = req.body.default_email;
  // Get the current user's profile photo.
  graphHelper.getProfilePhoto(accessToken, (errPhoto, profilePhoto) => {
    // //// TODO: MSA flow with local file (using fs and path?)
    if (!errPhoto) {
        // Upload profile photo as file to OneDrive.
        graphHelper.uploadFile(accessToken, profilePhoto, (errFile, file) => {
          // Get sharingLink for file.
          graphHelper.getSharingLink(accessToken, file.id, (errLink, link) => {
            const mailBody = emailer.generateMailBody(
              displayName,
              destinationEmailAddress,
              link.webUrl,
              profilePhoto
            );
            callback(null, mailBody);
          });
        });
      }
      else {
        var fs = require('fs');
        var readableStream = fs.createReadStream('public/img/test.jpg');
        var picFile;
        var chunk;
        readableStream.on('readable', function() {
          while ((chunk=readableStream.read()) != null) {
            picFile = chunk;
          }
      });
      
      readableStream.on('end', function() {

        graphHelper.uploadFile(accessToken, picFile, (errFile, file) => {
          // Get sharingLink for file.
          graphHelper.getSharingLink(accessToken, file.id, (errLink, link) => {
            const mailBody = emailer.generateMailBody(
              displayName,
              destinationEmailAddress,
              link.webUrl,
              picFile
            );
            callback(null, mailBody);
          });
        });
      });
      }
  });
}
*/
// Send an email.

/*
router.post('/sendMail', (req, res) => {
  const response = res;
  const templateData = {
    display_name: req.user.profile.displayName,
    email_address: req.user.profile.emails[0].address,
    actual_recipient: req.body.default_email
  };
  prepForEmailMessage(req, (errMailBody, mailBody) => {
    if (errMailBody) renderError(errMailBody);
    graphHelper.postSendMail(req.user.accessToken, JSON.stringify(mailBody), (errSendMail) => {
      if (!errSendMail) {
        response.render('sendMail', templateData);
      } else {
        if (hasAccessTokenExpired(errSendMail)) {
          errSendMail.message += ' Expired token. Please sign out and sign in again.';
        }
        renderError(errSendMail, response);
      }
    });
  });
});*/
