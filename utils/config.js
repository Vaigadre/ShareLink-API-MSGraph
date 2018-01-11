/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

module.exports = {
  creds: {
    redirectUrl: 'http://localhost:4000/token',
    clientID: '613b8f1d-fc68-4795-a2a5-2b558b7cb442',
    clientSecret: 'dLpbkQtSBqIB0PWk+yhEfIJ3ciW2bZIy0SRoFK7B4Ls=',
    identityMetadata: 'https://login.microsoftonline.com/triconinfotech.com/v2.0/.well-known/openid-configuration',
    allowHttpForRedirectUrl: true, // For development only
    responseType: 'code',
    validateIssuer: false, // For development only
    responseMode: 'query',
    scope: ['User.Read', 'Mail.Send', 'Files.ReadWrite']
  }
};
