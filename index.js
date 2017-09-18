var server = require('./server');
var router = require('./router');
var authHelper = require('./authHelper');
var url = require('url');
var microsoftGraph = require("@microsoft/microsoft-graph-client");

var handle = {};
handle['/'] = home;
handle['/authorize'] = authorize;

server.start(router.route, handle);

function home(response, request) {
    console.log('Request handler \'home\' was called.');
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write('<p>Please <a href="' + authHelper.getAuthUrl() + '">sign in</a> with your Office 365 or Outlook.com account.</p>');
    response.end();
}

function authorize(response, request) {
    console.log('Request handler \'authorize\' was called.');

    // The authorization code is passed as a query parameter
    var url_parts = url.parse(request.url, true);
    var code = url_parts.query.code;
    console.log('Code: ' + code);
    authHelper.getTokenFromCode(code, tokenReceived, response);
}

function tokenReceived(response, error, token) {
    if (error) {
        console.log('Access token error: ', error.message);
        response.writeHead(200, {'Content-Type': 'text/html'});
        response.write('<p>ERROR: ' + error + '</p>');
        response.end();
    } else {
        getUserEmail(token.token.access_token, function(error, email) {
            if (error) {
                console.log('getUserEmail returned an error: ' + error);
                response.write('<p>ERROR: ' + error + '</p>');
                response.end();
            } else if (email) {
                response.writeHead(200, {'Content-Type': 'text/html'});
                response.write('<p>Email: ' + email + '</p>');
                response.write('<p>Access token: ' + token.token.access_token + '</p>');
                response.end();
            }
        });
    }
}

function getUserEmail(token, callback) {
    // Create a Graph client
    var client = microsoftGraph.Client.init({
            authProvider: (done) => {
            // Just return the token
            done(null, token);
}
});

    // Get the Graph /Me endpoint to get user email address
    client
        .api('/me')
        .get((err, res) => {
        if (err) {
            callback(err, null);
        } else {
            callback(null, res.mail);
}
});
}