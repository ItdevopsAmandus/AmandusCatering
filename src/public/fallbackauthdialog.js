// MSAL authentication logic extracted from fallbackauthdialog.html
// Using Authorization Code Flow with PKCE

const msalConfig = {
  auth: {
    clientId: "82d99688-d922-4bfc-8d2d-e2871eb05ebd", // CLIENT_ID
    authority: "https://login.microsoftonline.com/82022306-deb0-41be-94c4-763bf46d3547",
    redirectUri: window.location.origin + "/fallbackauthdialog.html",
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: "localStorage",
  },
  system: {
    // Required for redirect login inside Office dialog
    allowRedirectInIframe: true,
  },
};

const loginRequest = {
  scopes: ["User.Read", "Sites.ReadWrite.All"],
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

function sendResult(status, payload) {
  Office.onReady(() => {
    Office.context.ui.messageParent(
      JSON.stringify({
        status: status,
        ...(status === "success" ? { result: payload } : { error: payload }),
      })
    );
  });
}

function acquireToken(account) {
  msalInstance
    .acquireTokenSilent({ ...loginRequest, account })
    .then((response) => {
      sendResult("success", response.accessToken);
    })
    .catch(() => {
      msalInstance.acquireTokenRedirect(loginRequest);
    });
}

msalInstance
  .handleRedirectPromise()
  .then((response) => {
    if (response !== null && response.account) {
      acquireToken(response.account);
    } else {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        acquireToken(accounts[0]);
      } else {
        msalInstance.loginRedirect(loginRequest);
      }
    }
  })
  .catch((error) => {
    console.error("MSAL-fout:", error);
    sendResult("failure", error.errorMessage || error.message || error);
  });
