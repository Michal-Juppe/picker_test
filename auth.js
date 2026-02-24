const msalParams = {
    auth: {
        authority: "https://login.microsoftonline.com/8820241f-63cd-4a58-8dd7-b2a9829ae87b",
        clientId: "7a4d1385-9b7c-47d3-8669-1abf187d7483",
        redirectUri: "https://michal-juppe.github.io/picker_test/"
    },
}

const app = new msal.PublicClientApplication(msalParams);

async function getToken(command) {

    let accessToken = "";
    let authParams = null;

    switch (command.type) {
        case "SharePoint":
        case "SharePoint_SelfIssued":
            authParams = { scopes: [`${combine(command.resource, ".default")}`] };
            break;
        default:
            break;
    }

    try {

        // see if we have already the idtoken saved
        const resp = await app.acquireTokenSilent(authParams);
        accessToken = resp.accessToken;

    } catch (e) {

        // per examples we fall back to popup
        const resp = await app.loginPopup(authParams);
        app.setActiveAccount(resp.account);

        if (resp.idToken) {

            const resp2 = await app.acquireTokenSilent(authParams);
            accessToken = resp2.accessToken;

        }
    }

    return accessToken;
}
