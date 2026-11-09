import { useState } from 'react';
import { type DriveFile } from '@glennjong/vibe-sheets';

import type { Data } from '../types';
import { fetchScript } from '../common/fetch';


const FileItem = ({ isNew, file, onSelect }: { isNew: boolean, file: DriveFile, onSelect: (data: Data[]) => void }) => {
  
  const [ isFetching, setIsFetching ] = useState(false);
  const [ isAccessed, setIsAccessed ] = useState<boolean | undefined>(!isNew);
  
  const handleSelect = async () => {
    try {
      if (file.description) {
        if (file.scriptUrl) {
          setIsFetching(true);
          const data = await fetchScript(file.scriptUrl);
          setIsAccessed(true);
          onSelect(data);
          setIsFetching(false);
        }
      }
    } catch (e) {
      console.error("Invalid file description", e);
      setIsAccessed(false);
      setIsFetching(false);
    }
  };

  // View: List Files
  return (
    <div style={{ 
      backgroundColor: 'var(--bg-card)', 
      border: '1px solid var(--border-color)', 
      color: 'var(--text-main)',
      borderRadius: '8px',
      padding: '16px',
      marginBottom: '12px',
      boxShadow: '0 2px 5px var(--shadow-color)',
      transition: 'transform 0.1s ease-in-out',
    }}>
      <div key={file.id}>
        <div style={{ display: 'flex', alignItems: 'center', marginBottom: '12px' }}>
          <span className="material-icons" style={{ fontSize: '24px', marginRight: '8px', color: 'var(--primary)' }}>
            description
          </span>
          <h3 style={{ margin: 0, fontSize: '1.1rem', fontWeight: 'bold' }}>{file.name.replace('moneybook-', '')}</h3>
        </div>
        
        <div style={{ display: 'flex', gap: '8px' }}>
          <a href={file.webViewLink} target="_blank" rel="noreferrer" style={{ flex: 1, textDecoration: 'none' }}>
            <button style={{ 
              width: '100%', 
              padding: '10px',
              background: 'var(--bg-item)', 
              color: 'var(--text-main)', 
              border: '1px solid var(--border-color)',
              borderRadius: '6px',
              cursor: 'pointer',
              fontWeight: '500',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              gap: '4px'
            }}>
              <span className="material-icons" style={{ fontSize: '16px' }}>open_in_new</span>
              Open
            </button>
          </a>
          <button 
            disabled={isFetching || isAccessed !== true}
            onClick={handleSelect}
            style={{ 
              flex: 2,
              padding: '10px',
              backgroundColor: (isFetching || isAccessed !== true) ? 'var(--input-bg)' : 'var(--primary)',
              color: (isFetching || isAccessed !== true) ? 'var(--text-muted)' : 'var(--text-inv)',
              border: 'none',
              borderRadius: '6px',
              cursor: (isFetching || isAccessed !== true) ? 'not-allowed' : 'pointer',
              fontWeight: 'bold',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              gap: '4px'
            }}
          >
            { isFetching ? (
              <>
                 <span className="spinner" style={{ width: '14px', height: '14px', border: '2px solid', borderTopColor: 'transparent', borderRadius: '50%', display: 'inline-block', animation: 'spin 1s linear infinite' }}></span>
                 Loading
                 <style>{`@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }`}</style>
              </>
            ) : (
              <>
                <span className="material-icons" style={{ fontSize: '18px' }}>check</span>
                Select
              </>
            )}
          </button>
        </div>

        { isAccessed === false && 
          <div style={{ marginTop: '12px', padding: '10px', backgroundColor: 'var(--warning-bg-subtle)', borderRadius: '6px', border: '1px solid var(--warning)' }}>
            <p style={{ margin: '0 0 8px 0', color: 'var(--text-main)', fontSize: '0.9rem' }}>Access authorization required.</p>
            <button 
              onClick={() => {
                const scriptUrl = file.scriptUrl;
                window.open(`${scriptUrl}?action=authorize`, 'auth', 'width=600,height=400');
                setIsAccessed(true);
              }}
              style={{
                width: '100%',
                backgroundColor: 'var(--primary)',
                color: 'var(--text-inv)',
                border: 'none',
                padding: '8px',
                borderRadius: '4px',
                cursor: 'pointer',
                fontWeight: '500'
              }}
            >
              Authorize
            </button>
          </div>
        }
      </div>
    </div>
  );
};

export default FileItem;
