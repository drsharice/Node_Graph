// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

var graph = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

module.exports = {
  getUserDetails: async function(msalClient, userId) {
    const client = getAuthenticatedClient(msalClient, userId);

    const user = await client
      .api('/me')
      .select('displayName,mail,mailboxSettings,userPrincipalName')
      .get();
    return user;
  },

  // <GetCalendarViewSnippet>
  getCalendarView: async function(msalClient, userId, start, end, timeZone) {
    const client = getAuthenticatedClient(msalClient, userId);

    const events = await client
      .api('/me/calendarview')
      // Add Prefer header to get back times in user's timezone
      .header("Prefer", `outlook.timezone="${timeZone}"`)
      // Add the begin and end of the calendar window
      .query({ startDateTime: start, endDateTime: end })
      // Get just the properties used by the app
      .select('subject,organizer,start,end')
      // Order by start time
      .orderby('start/dateTime')
      // Get at most 50 results
      .top(50)
      .get();

    return events;
  },
  // </GetCalendarViewSnippet>

  // <CreateEventSnippet>
  createEvent: async function(msalClient, userId, formData, timeZone) {
    const client = getAuthenticatedClient(msalClient, userId);

    // Build a Graph event
    const newEvent = {
      subject: formData.subject,
      start: {
        dateTime: formData.start,
        timeZone: timeZone
      },
      end: {
        dateTime: formData.end,
        timeZone: timeZone
      },
      body: {
        contentType: 'text',
        content: formData.body
      }
    };

    // Add attendees if present
    if (formData.attendees) {
      newEvent.attendees = [];
      formData.attendees.forEach(attendee => {
        newEvent.attendees.push({
          type: 'required',
          emailAddress: {
            address: attendee
          }
        });
      });
    }

    // POST /me/events
    await client
      .api('/me/events')
      .post(newEvent);
  },
  // </CreateEventSnippet>
    // <get email>
    getMailEvents: async function (msalClient) {
        console.log('Hello');
        const client = getAuthenticatedClient(msalClient);
        try {
            let messages = await client
                .api('/me/messages')
                .select('subject,receivedDateTime,bodyPreview,sender')
                .orderby('receivedDateTime')
                .top(100)
                .get();

            return messages;
        }
        catch (err) {
            console.log(JSON.stringify(err));
        }

    },

    // <CreateMessageSnippet>
    createMessage: async function (msalClient, formData, timeZone) {
        const client = getAuthenticatedClient(msalClient);

        // Build a Graph event
        const newMessage = {
            subject: formData.subject,
            body: {
                contentType: 'text',
                content: formData.body
            },
            toRecipients: {
                contentType: 'text',
                content: formData.toRecipients
            }

        };
        // POST /me/events
        await client
            .api('/me/message')
            .post(newMessage);
    },
    // <SendMessageSnippet>
    sendMessage: async function (msalClient, formData, timeZone) {
        const client = getAuthenticatedClient(msalClient);

        // send message
        const sendMessage = {
            subject: formData.subject,
            body: {
                contentType: 'text',
                content: formData.body
            },
            toRecipients: {
                contentType: 'text',
                content: formData.toRecipients
            }

        };
        // POST /me/events
        await client
            .api('/me/sendMail')
            .post(sendMessage);
    },


    getJoinTeamEvents: async function (msalClient) {

        const client = getAuthenticatedClient(msalClient);
        try {
            let joinedTeams = await client.api('/me/joinedTeams')
                .get();

            return joinedTeams;
        }
        catch (err) {
            console.log(JSON.stringify(err));
        }

    },
    //Not being used at the moment
    getEmails: async function () {
        ensureScope('mail.read');

        return await graphClient
            .api('/me/messages')
            .select('subject,receivedDateTime,bodyPreview')
            .orderby('receivedDateTime desc')
            .top(10)
            .get();
    }

};

function getAuthenticatedClient(msalClient, userId) {
  if (!msalClient || !userId) {
    throw new Error(
      `Invalid MSAL state. Client: ${msalClient ? 'present' : 'missing'}, User ID: ${userId ? 'present' : 'missing'}`);
  }

  // Initialize Graph client
  const client = graph.Client.init({
    // Implement an auth provider that gets a token
    // from the app's MSAL instance
    authProvider: async (done) => {
      try {
        // Get the user's account
        const account = await msalClient
          .getTokenCache()
          .getAccountByHomeId(userId);

        if (account) {
          // Attempt to get the token silently
          // This method uses the token cache and
          // refreshes expired tokens as needed
          const response = await msalClient.acquireTokenSilent({
            scopes: process.env.OAUTH_SCOPES.split(','),
            redirectUri: process.env.OAUTH_REDIRECT_URI,
            account: account
          });

          // First param to callback is the error,
          // Set to null in success case
          done(null, response.accessToken);
        }
      } catch (err) {
        console.log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
        done(err, null);
      }
    }
  });

  return client;
}
