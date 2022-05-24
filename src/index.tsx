import React from 'react';
import ReactDOM from 'react-dom/client';
import './index.css';
import App from './App';
import reportWebVitals from './reportWebVitals';
import { AuthenticationResult, Configuration, EventMessage, EventType, PublicClientApplication } from '@azure/msal-browser';
import { MsalProvider } from '@azure/msal-react';

const redirectUri = () => {
  const location = document.location;
  if (location.hostname === 'localhost') {
    console.log(`${location.protocol}//${location.hostname}:${location.port}${location.pathname}`)
    return `${location.protocol}//${location.hostname}:${location.port}${location.pathname}`
  }
  
  return `${location.protocol}//${location.hostname}${location.pathname}`
}
const configuration: Configuration = {
  auth: {
    clientId: '69d4c8a4-52c4-478f-9361-b7860e8c7018',
    authority: 'https://login.microsoftonline.com/7d13004c-123c-4064-9291-3ebefd2822df',
    redirectUri: redirectUri()
  },
}

const pca = new PublicClientApplication(configuration);

const accounts = pca.getAllAccounts();
if (accounts.length > 0) {
    pca.setActiveAccount(accounts[0]);
}
pca.addEventCallback((event: EventMessage) => {
  if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
      const payload = event.payload as AuthenticationResult;
      const account = payload.account;
      pca.setActiveAccount(account);
  }
});

const root = ReactDOM.createRoot(
  document.getElementById('root') as HTMLElement
);
root.render(
  <MsalProvider instance={pca}>
    <React.StrictMode>
      <App />
    </React.StrictMode>
  </MsalProvider>
);

// If you want to start measuring performance in your app, pass a function
// to log results (for example: reportWebVitals(console.log))
// or send to an analytics endpoint. Learn more: https://bit.ly/CRA-vitals
reportWebVitals();
