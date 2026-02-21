import { useState, lazy, Suspense } from 'react';
import './App.css'
import MainLayout from './screen/MainLayout';

const SetupScreen = lazy(() => import('./screen/SetupScreen'));

function App() {
  const [selectedScriptUrl, setSelectedScriptUrl] = useState<string | null>(
    localStorage.getItem('vibe_script_url') || null
  );

  if (selectedScriptUrl) {
    return (
      <MainLayout />
    );
  }
  
  return (
    <Suspense fallback={
        <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', height: '100vh' }}>
            <span className="spinner" style={{ width: '30px', height: '30px', border: '3px solid var(--primary)', borderTopColor: 'transparent', borderRadius: '50%', animation: 'spin 1s linear infinite' }}></span>
        </div>
    }>
      <SetupScreen onScriptSelected={setSelectedScriptUrl} />
    </Suspense>
  )
}

export default App

