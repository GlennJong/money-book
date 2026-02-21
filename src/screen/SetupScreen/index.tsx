import { lazy, Suspense } from 'react';
import { ReactHooks } from '@glennjong/vibe-sheets';
import '../../App.css'

const SheetSelector = lazy(() => import('../../SheetSelector'));

const clientId = import.meta.env.VITE_GOOGLE_CLIENT_ID;

const SetupScreen = ({
  onScriptSelected
}: {
  onScriptSelected: (url: string) => void
}) => {
  const { login, accessToken, isAppsScriptEnabled } = ReactHooks.useGoogleAuth({
    clientId
  });

  // Handle getting GAS permission if it's not enabled
  if (!accessToken) {
    return (
      <div className="card" style={{ color: 'var(--text-main)', textAlign: 'center', padding: '40px 20px' }}>
        <img src="icons/icon-192x192.png" alt="MoneyBook Logo" style={{ width: '100px', height: '100px', borderRadius: '20px', marginBottom: '20px', boxShadow: '0 4px 10px var(--shadow-color)' }} />
        <h1 style={{ fontSize: '1.8em', marginBottom: '30px' }}>Welcome to MoneyBook</h1>
        <button onClick={login} style={{ fontSize: '1.1em', padding: '12px 24px', backgroundColor: 'var(--primary)', color: 'var(--bg-card)', border: 'none', borderRadius: '8px', cursor: 'pointer' }}>Login with Google</button>
      </div>
    );
  } else if (accessToken && !isAppsScriptEnabled) {
    return (
      <div className="card" style={{ backgroundColor: 'var(--bg-card)', color: 'var(--text-main)', border: '1px solid var(--border-color)', borderRadius: '8px' }}>
        {isAppsScriptEnabled === undefined ?
          <p>Checking your Apps Script permission...</p>
          :
          <>
            <p>Before we start using MoneyBook, you need to enable Apps Script.</p>
            <a target="_blank" href="https://script.google.com/home/usersettings?pli=1">Click Me</a>
          </>
        }
      </div>
    );
  }

  return (
    <Suspense fallback={
        <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100vh' }}>
            <span className="spinner" style={{ width: '30px', height: '30px', border: '3px solid var(--primary)', borderTopColor: 'transparent', borderRadius: '50%', animation: 'spin 1s linear infinite' }}></span>
        </div>
    }>
      <SheetSelector
        token={accessToken}
        onSelect={(endpoint: string) => {
          localStorage.setItem('vibe_script_url', endpoint);
          onScriptSelected(endpoint);
        }}
      />
    </Suspense>
  )
};

export default SetupScreen;
