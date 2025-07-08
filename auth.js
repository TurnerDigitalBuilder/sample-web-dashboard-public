// Configuration for the Microsoft Authentication Library (MSAL).
// These values come from your Azure AD application registration.
const msalConfig = {
    auth: {
        // TODO: Replace with your Application (client) ID from Azure AD App Registration
        // Found in: Azure Portal -> Azure Active Directory -> App registrations -> Your App -> Overview
        clientId: "YOUR_CLIENT_ID_HERE", 
        
        // TODO: Replace with your Directory (tenant) ID 
        // Found in: Azure Portal -> Azure Active Directory -> Overview -> Tenant ID
        // Format: https://login.microsoftonline.com/YOUR_TENANT_ID_HERE
        authority: "https://login.microsoftonline.com/YOUR_TENANT_ID_HERE", 
        
        // TODO: Replace with your application's redirect URI
        // This should match what you configured in Azure AD App Registration -> Authentication -> Redirect URIs
        // For local development: "http://localhost:3000" or "http://localhost:8080"
        // For production: "https://your-domain.com" or "https://your-github-username.github.io/your-repo-name"
        redirectUri: "YOUR_REDIRECT_URI_HERE"
    },
    cache: {
        // Store tokens in session storage so they clear when the tab closes
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    }
};

// Create an authentication instance which manages the sign-in process
const msalInstance = new msal.PublicClientApplication(msalConfig);

// When returning from a redirect-based login MSAL provides the token here
msalInstance.handleRedirectPromise()
    .then(tokenResponse => {
        if (tokenResponse) {
            sessionStorage.setItem("msalAccount", tokenResponse.account.username);
            updateUI();
        }
    }).catch(error => {
        console.error(error);
    });

function signIn() {
    // Ask the user to sign in. Scopes define what permissions we request.
    const loginRequest = {
        scopes: ["User.Read", "Sites.ReadWrite.All"]
    };
    msalInstance.loginRedirect(loginRequest);
}

function signOut() {
    // Sign the current user out and clear our session state
    const logoutRequest = {
        account: msalInstance.getAccountByUsername(sessionStorage.getItem("msalAccount"))
    };
    msalInstance.logoutRedirect(logoutRequest);
    sessionStorage.removeItem("msalAccount");
}

async function getAccessToken() {
    // Look up the currently signed-in user
    const account = msalInstance.getAccountByUsername(sessionStorage.getItem("msalAccount"));
    if (!account) {
        throw new Error("User not signed in");
    }
    const tokenRequest = {
        scopes: ["User.Read", "Sites.ReadWrite.All"],
        account: account
    };
    try {
        // Try to get a token without showing any UI
        const tokenResponse = await msalInstance.acquireTokenSilent(tokenRequest);
        return tokenResponse.accessToken;
    } catch (err) {
        if (err instanceof msal.InteractionRequiredAuthError) {
            // If silent auth fails we redirect the user for an interactive sign-in
            return msalInstance.acquireTokenRedirect(tokenRequest);
        }
        console.error(err);
    }
}