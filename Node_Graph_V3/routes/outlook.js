const router = require('express-promise-router')();
const graph = require('../graph.js');
const moment = require('moment-timezone');
const iana = require('windows-iana');
const { body, validationResult } = require('express-validator');
const validator = require('validator');

/* GET /outlook*/
router.get('/',
    async function (req, res) {
        if (!req.session.userId) {
            // Redirect unauthenticated requests to home page
            res.redirect('/')
        } else {
            const params = {
                active: { outlook: true }
            };

          
            // Get the access token
            var accessToken;
            try {
                accessToken = await getAccessToken(req.session.userId, req.app.locals.msalClient);
            } catch (err) {
                res.send(JSON.stringify(err, Object.getOwnPropertyNames(err)));
                return;
            }

            if (accessToken && accessToken.length > 0) {
                try {
                    // Get the events
                    const mailevents = await graph.getMailEvents(accessToken);                    
                    res.json(mailevents.value);
                } catch (err) {
                    res.send(JSON.stringify(err, Object.getOwnPropertyNames(err)));
                }
            }
            else {
                req.flash('error_msg', 'Could not get an access token');
            }
        }
    }
);

async function getAccessToken(userId, msalClient) {
    // Look up the user's account in the cache
    try {
        const accounts = await msalClient
            .getTokenCache()
            .getAllAccounts();

        const userAccount = accounts.find(a => a.homeAccountId === userId);

        // Get the token silently
        const response = await msalClient.acquireTokenSilent({
            scopes: process.env.OAUTH_SCOPES.split(','),
            redirectUri: process.env.OAUTH_REDIRECT_URI,
            account: userAccount
        });

        return response.accessToken;
    } catch (err) {
        console.log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
    }
}

module.exports = router;