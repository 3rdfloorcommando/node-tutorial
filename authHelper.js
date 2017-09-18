var credentials = {
    client: {
        id: '753e5b53-a75b-46bc-8fcf-e0331186806b',
        secret: 'LrPFFayR3VVfDse0S3TNPmD',
    },
    auth: {
        tokenHost: 'https://login.microsoftonline.com',
        authorizePath: 'common/oauth2/v2.0/authorize',
        tokenPath: 'common/oauth2/v2.0/token'
    }
};
var oauth2 = require('simple-oauth2').create(credentials);

var redirectUri = 'http://localhost:8000/authorize';

// The scopes the app requires
var scopes = [ 'openid',
    'https://outlook.office.com/mail.read','https://outlook.office.com/calendars.readwrite' ];

function getAuthUrl() {
    var returnVal = oauth2.authorizationCode.authorizeURL({
        redirect_uri: redirectUri,
        scope: scopes.join(' ')
    });
    console.log('Generated auth url: ' + returnVal);
    return returnVal;
}

function getTokenFromCode(auth_code, callback, response) {
    var token;
    oauth2.authorizationCode.getToken({
        code: auth_code,
        redirect_uri: redirectUri,
        scope: scopes.join(' ')
    }, function (error, result) {
        if (error) {
            console.log('Access token error: ', error.message);
            callback(response, error, null);
        } else {
            token = oauth2.accessToken.create(result);
            console.log('Token created: ', token.token);
            callback(response, null, token);
        }
    });
}

exports.getTokenFromCode = getTokenFromCode;
exports.getAuthUrl = getAuthUrl;