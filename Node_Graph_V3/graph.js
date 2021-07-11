// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

var graph = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');

module.exports = {
    getUserDetails: async function (accessToken) {
        const client = getAuthenticatedClient(accessToken);

        const user = await client
            .api('/me')
            .select('displayName,mail,mailboxSettings,userPrincipalName')
            .get();
        return user;
    },

    // <GetCalendarViewSnippet>
    getCalendarView: async function (accessToken, start, end, timeZone) {
        const client = getAuthenticatedClient(accessToken);

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
    createEvent: async function (accessToken, formData, timeZone) {
        const client = getAuthenticatedClient(accessToken);

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
    // <get email>
    getMailEvents: async function (accessToken) {
        console.log('Hello');
        const client = getAuthenticatedClient(accessToken);
        try {
            let messages = await client
                .api('/me/messages')
                .select('subject,receivedDateTime,bodyPreview,sender')
                .orderby('receivedDateTime desc')
                .top(50)
                .get();

            return messages;
        }
        catch (err) {
            console.log(JSON.stringify(err));
        }

    },
    getChatEvents: async function (accessToken) {

        const client = getAuthenticatedClient(accessToken);
        try {
            let chatmessages = await client.api('/teams/fbe2bf47-16c8-47cf-b4a5-4b9b187c508b/channels/19:4a95f7d8db4c4e7fae857bcebe0623e6@thread.tacv2/messages')
                .query('from,body')
                .top(2)
                .get();

            return chatmessages;
        }
        catch (err) {
            console.log(JSON.stringify(err));
        }

    },
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

function getAuthenticatedClient(accessToken) {
    // Initialize Graph client
    const client = graph.Client.init({
        // Use the provided access token to authenticate
        // requests
        authProvider: (done) => {
            done(null, accessToken);
        }
    });

    return client;
}
