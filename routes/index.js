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
  res.render('login');
   //res.redirect('/login');  
});

// Authentication request.
router.get('/login',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }),
    (req, res) => {
      res.redirect('/token');  
    });

// Authentication callback.
// After we have an access token, get user data and load the sendMail page.
router.get('/token',
  passport.authenticate('azuread-openidconnect', { failureRedirect: '/' }),
    (req, res) => {

      //res.setHeader("200", {'Content-Type': 'application/json'});
      //shareLink(req, res);

    graphHelper.getUserData(req.user.accessToken).then( (user) => {
        console.log(user.body.displayName);
    }).catch(err => {console.log(err)})

         //  req.user.profile.displayName = user.body.displayName;
          //  req.user.profile.emails = [{ address: user.body.mail || user.body.userPrincipalName }];
          // renderSendMail(req, res);
           shareLink (req, res);
          // res.send('<p>Operation completed<p>');
           
          // graphHelper.getDriveFileList(req.user.accessToken, (err, list) => {
          //   console.log(list)
          // })
  //       } else {
  //         renderError(err, res);
  //         res.send('<p>Operation failed<p>');
  //       }
  //   });
  // }*/
 }); 

function shareLink (req, res)  {
  const accessToken = req.user.accessToken;
   const user = {
    display_name: req.user.profile.displayName,
    accessToken: req.user.accessToken
  }

  console.log(req.query.session_state);

  graphHelper.getDriveFileList(accessToken)
   .then( file =>{
     return graphHelper.copyFileFromDrive(accessToken, file.value[0].id);
   }).then ( id => {
     return graphHelper.insertDataToExcel(accessToken, id);
   }).then ( (id) => {
     console.log('Response after insertion in sheet: \n' + id);
     return graphHelper.getSharingLink(accessToken, id);
   }).then( shareLink => {
    console.log('Copied file is available on URL: \n' + shareLink);
    open(shareLink);    
   }).catch( err => {
    // console.log(err);
   })

    // res.render('fileList', {
    //   user: user,
    //   link: 'www.google.com',
    //   actionUrl: "https://INC-excel.officeapps.live.com/x/_layouts/xlviewerinternal.aspx?edit=1&ui=en%2DUS&rs=en%2DUS&WOPISrc=https%3A%2F%2Ftriconindia%2Dmy%2Esharepoint%2Ecom%2Fpersonal%2Fvaibhav%5Ftriconinfotech%5Fcom%2F%5Fvti%5Fbin%2Fwopi%2Eashx%2Ffiles%2F56d94e388221479780233e0f51609c63&activeCell=%27Sheet1%27%21A1&wdInitialSession=f9e9a9d9%2D7582%2D46e2%2Db740%2Dcf3fae4857f3&wdRldC=1&wdEnableRoaming=1&mscc=1&wdODB=1", 
    //   accessToken: accessToken,
    //   access_token_ttl: (new Date).getTime()
    // })  
//   res.send(user);
};

  // console.log("action " + actionUrl);

    // res.render('fileList', {
    //   user: user,
    //   link: 'www.google.com',
    //   actionUrl: actionUrl, accessToken: accessToken
    // })
  
  //  f9e9a9d9-7582-46e2-b740-cf3fae4857f3 
   

  // graphHelper.getDriveFileList(accessToken, (err, file) => {
  //   if (err) { console.log({ err: err.message }) }

  //   console.log('Copying file With ID: ' + file.value[1].id)

  //   graphHelper.copyFileFromDrive(accessToken, file.value[1].id, (err, id) => {
  //     console.log('Copied new file ID: ' + id)

  //     graphHelper.insertDataToExcel(accessToken, id, (err, data) => {
  //       if (err) { console.log({ err: err.message }) }

  //       console.log('Response after insertion in sheet: \n' + data)

  //       graphHelper.getSharingLink(accessToken, id, (err, shareLink) => {
           
  //         if (err) { console.log({ err: err.message }) }
  //         open(shareLink)
  //         console.log('Copied file is available on URL: \n' + shareLink)

          //res.redirect(link)
          // res.render('fileList', {
          //   user: user,
          //   link: shareLink, 
          //   url: url
          // })

    //     })
    //   })      
    // })
  //})





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
