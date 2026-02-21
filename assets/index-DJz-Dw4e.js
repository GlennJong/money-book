const __vite__mapDeps=(i,m=__vite__mapDeps,d=(m.f||(m.f=["./SheetSelector-DWJA5MHS.js","./index-Erze52nw.js","./index-D_NmA3Ee.css"])))=>i.map(i=>d[i]);
import{r as C,j as E,_ as H}from"./index-Erze52nw.js";var W=Object.defineProperty,Y=(e,w)=>{for(var u in w)W(e,u,{get:w[u],enumerable:!0})},L={doGet:`
    function doGet(e) {
      try {
        var sheetName = e.parameter.sheet;
        var sheet;
        if (sheetName) {
          sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
          if (!sheet) {
             return createJsonResponse({ error: 'Sheet "' + sheetName + '" not found' });
          }
        } else {
          sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
        }

        var rows = sheet.getDataRange().getValues();
        
        if (rows.length === 0) {
          return createJsonResponse({ data: [] });
        }

        var headers = rows[0];
        
        // 解析 fields 參數 (例如: ?fields=name,value 或 ?fields=name+value)
        var fieldsParam = e.parameter.fields;
        var allowedFields = null;
        if (fieldsParam) {
           // 支援逗號或空格分隔
           allowedFields = fieldsParam.split(/[,s+]+/).filter(function(f) { return f && f.trim().length > 0; });
        }

        // 1. 先處理 Soft Delete 過濾
        var isEnabledIndex = headers.indexOf('is_enabled');
        var validRows = rows.slice(1);
        
        if (isEnabledIndex !== -1) {
             validRows = validRows.filter(function(row) {
                var check = row[isEnabledIndex];
                return check !== false && check !== 'FALSE';
             });
        }
        
        // 2. 再處理欄位 Mapping 與 Filter
        var data = validRows.map(function(row) {
          var obj = {};
          headers.forEach(function(header, i) {
            if (header === 'is_enabled') return; // Always exclude is_enabled from output

            if (allowedFields) {
                // 修改：即使 filter 中沒有指定 id，也強制回傳 id 供前端辨識
                if (header === 'id' || allowedFields.indexOf(header) !== -1) {
                    obj[header] = row[i];
                }
            } else {
                obj[header] = row[i];
            }
          });
          return obj;
        });

        return createJsonResponse({ data: data });
      } catch (err) {
        return createJsonResponse({ error: err.toString() });
      }
    }
  `,doPost:`
    function doPost(e) {
      try {
        var sheetName = e.parameter.sheet;
        var sheet;
        if (sheetName) {
          sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
          if (!sheet) {
             return createJsonResponse({ error: 'Sheet "' + sheetName + '" not found' });
          }
        } else {
          sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
        }

        var action = e.parameter.method || e.parameter.action;

        // --- UPDATE Logic ---
        if (action === 'PUT' || action === 'UPDATE') {
           var updateData;
           try {
              updateData = JSON.parse(e.postData.contents);
           } catch (err) {
              return createJsonResponse({ error: 'Invalid JSON for update', debug: err.toString() });
           }
           
           if (!updateData.id) {
               return createJsonResponse({ error: 'Update requires an "id" field' });
           }

           var lastRow = sheet.getLastRow();
           var lastCol = sheet.getLastColumn();
           if (lastRow <= 1) return createJsonResponse({ error: 'No data to update' });
           
           var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
           var idIndex = headers.indexOf('id');
           
           if (idIndex === -1) return createJsonResponse({ error: 'Sheet needs an "id" column' });
           
           // Find Row by ID
           var allIds = sheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues().map(function(r) { return r[0]; });
           var rowIndex = -1;
           var targetId = String(updateData.id);
           
           for (var i = 0; i < allIds.length; i++) {
               if (String(allIds[i]) === targetId) {
                   rowIndex = i + 2; 
                   break;
               }
           }
           
           if (rowIndex === -1) {
               return createJsonResponse({ error: 'ID not found: ' + updateData.id });
           }
           
           // Update Fields
           var updatedFields = [];
           Object.keys(updateData).forEach(function(key) {
               if (key === 'id') return; // Don't update ID
               
               var colIndex = headers.indexOf(key);
               if (colIndex !== -1) {
                   sheet.getRange(rowIndex, colIndex + 1).setValue(updateData[key]);
                   updatedFields.push(key);
               }
           });
           
           // Update updated_at
           var updatedAtIndex = headers.indexOf('updated_at');
           if (updatedAtIndex !== -1) {
               sheet.getRange(rowIndex, updatedAtIndex + 1).setValue(new Date().toISOString());
           }
           
           return createJsonResponse({ 
               status: 'success', 
               message: 'Row updated', 
               updatedFields: updatedFields,
               id: updateData.id 
           });
        }

        // --- DELETE Logic (Soft Delete) ---
        if (action === 'DELETE') {
           var deleteData;
           try {
              deleteData = JSON.parse(e.postData.contents);
           } catch (err) {
              return createJsonResponse({ error: 'Invalid JSON for delete', debug: err.toString() });
           }
           
           if (!deleteData.id) {
               return createJsonResponse({ error: 'Delete requires an "id" field' });
           }

           var lastRow = sheet.getLastRow();
           var lastCol = sheet.getLastColumn();
           if (lastRow <= 1) return createJsonResponse({ error: 'No data to delete' });
           
           var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
           var idIndex = headers.indexOf('id');
           var enabledIndex = headers.indexOf('is_enabled');
           
           if (idIndex === -1) return createJsonResponse({ error: 'Sheet needs an "id" column' });
           // 若沒有 is_enabled 欄位，就無法做軟刪除，這邊視為錯誤或直接 return
           if (enabledIndex === -1) return createJsonResponse({ error: 'Sheet needs an "is_enabled" column for soft delete' });
           
           var allIds = sheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues().map(function(r) { return r[0]; });
           
           var targetId = String(deleteData.id);
           var rowIndex = -1;
           
           for (var i = 0; i < allIds.length; i++) {
               // 使用 String() 確保型別一致，避免 123 != "123" 的問題
               if (String(allIds[i]) === targetId) {
                   rowIndex = i + 2; 
                   break;
               }
           }
           
           if (rowIndex === -1) {
               return createJsonResponse({ error: 'ID not found: ' + deleteData.id });
           }
           
           // update is_enabled to false
           sheet.getRange(rowIndex, enabledIndex + 1).setValue(false);
           
           // update updated_at if exists
           var updatedAtIndex = headers.indexOf('updated_at');
           if (updatedAtIndex !== -1) {
               sheet.getRange(rowIndex, updatedAtIndex + 1).setValue(new Date().toISOString());
           }
           
           return createJsonResponse({ status: 'success', message: 'Row soft deleted (is_enabled=false)', id: deleteData.id });
        }

        // --- CREATE Logic (Default) ---
        
        // 1. 解析傳入的 JSON 資料
        var postData;
        try {
          postData = JSON.parse(e.postData.contents);
        } catch (err) {
           return createJsonResponse({ error: 'Invalid JSON', debug: err.toString() });
        }
        
        // 判斷是單筆還是多筆
        var incomingRows = [];
        if (Array.isArray(postData)) {
           incomingRows = postData;
        } else {
           incomingRows = [postData];
        }

        if (incomingRows.length === 0) {
            return createJsonResponse({ message: 'No data to insert' });
        }
        
        // 2. 獲取現有的標頭
        // 我們假設第一列是 headers
        var lastCol = sheet.getLastColumn();
        if (lastCol === 0) {
           return createJsonResponse({ error: 'Sheet is empty (no headers)' });
        }
        
        // 讀取第一列 (Header)
        var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        
        // 3. 準備寫入的資料
        var newRows = [];
        var timestamp = new Date();
        var createdIds = [];
        
        incomingRows.forEach(function(rowObj) {
            var newRow = [];
            headers.forEach(function(header) {
                var value = rowObj[header];
                
                // 自動填入欄位處理
                if (header === 'id') {
                    if (!value) {
                       // 簡單產生唯一 ID
                       value = Utilities.getUuid();
                    } else {
                       // 強制轉為字串
                       value = String(value);
                    }
                } else if (header === 'created_at' && !value) {
                    value = timestamp.toISOString();
                } else if (header === 'updated_at' && !value) {
                    value = timestamp.toISOString();
                } else if (header === 'is_enabled' && (value === undefined || value === "")) {
                    value = true;
                }
                
                // 沒有值就填空字串，避免 undefined
                if (value === undefined || value === null) {
                    value = "";
                }
                
                newRow.push(value);
            });
            newRows.push(newRow);

            // 假設 headers 裡有 'id'，收集起來回傳
            var idIndex = headers.indexOf('id');
            if (idIndex !== -1) {
                createdIds.push(newRow[idIndex]);
            }
        });
        
        // 4. 寫入 Spreadsheet
        if (newRows.length > 0) {
            // 修正：不直接使用 sheet.getLastRow()，因為如果有整列 checkbox，getLastRow 會回傳 maxRows
            // 我們改為偵測 'id' 欄位來決定真正的最後一行
            var lastRowWithData = 1; // 至少有 Header
            var idIndex = headers.indexOf('id');
            
            if (idIndex !== -1) {
                // 讀取整欄 ID (假設資料量不超過 Sheet 上限，若很多可分批讀取或用其他方式優化)
                // getRange(row, col, numRows)
                // 讀取從第2列開始的所有 ID
                var maxRows = sheet.getMaxRows();
                if (maxRows > 1) {
                    var idValues = sheet.getRange(2, idIndex + 1, maxRows - 1, 1).getValues();
                    // 由後往前找第一個有值的
                    for (var i = idValues.length - 1; i >= 0; i--) {
                        if (idValues[i][0] && idValues[i][0] !== "") {
                            lastRowWithData = i + 2; // array index + 2 (because started from row 2)
                            break;
                        }
                    }
                }
            } else {
                // Fallback (若無 ID 欄位)
                lastRowWithData = sheet.getLastRow();
            }

            var startRow = lastRowWithData + 1;
            
            // data: newRows is already prepared
            var startRow = lastRowWithData + 1;
            
            // setValues 需要二維陣列，且大小需完全符合 Range
            sheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
            
            // 動態為新增的資料列設定 Checkbox 驗證
            // 針對 is_enabled 欄位，或任何值為 boolean 的欄位自動套用 Checkbox
            headers.forEach(function(header, idx) {
                var isBooleanCol = (header === 'is_enabled');
                
                // 若不是 is_enabled，檢查首筆資料該欄位是否為 boolean (簡易自動判斷)
                if (!isBooleanCol && newRows.length > 0) {
                    var sampleVal = newRows[0][idx];
                    if (typeof sampleVal === 'boolean') {
                        isBooleanCol = true;
                    }
                }

                if (isBooleanCol) {
                    var range = sheet.getRange(startRow, idx + 1, newRows.length, 1);
                    var rule = SpreadsheetApp.newDataValidation()
                      .requireCheckbox()
                      .setAllowInvalid(false)
                      .build();
                    range.setDataValidation(rule);
                }
            });
        }
        
        // 5. 回傳成功訊息
        return createJsonResponse({ 
           status: 'success', 
           message: newRows.length + ' row(s) appended',
           createdIds: createdIds
        });
        
      } catch (err) {
        return createJsonResponse({ error: err.toString() });
      }
    }
  `,helpers:`
    function createJsonResponse(data) {
      return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
    }
  `},K=()=>[L.doGet,L.doPost,L.helpers].join(`

`),O={SHEETS:"https://sheets.googleapis.com/v4/spreadsheets",SCRIPT_PROJECTS:"https://script.googleapis.com/v1/projects",DRIVE_FILES:"https://www.googleapis.com/drive/v3/files"},U=e=>({Authorization:`Bearer ${e}`,"Content-Type":"application/json"});async function Z(e,w,u=[],p="default"){const m=[{name:"is_enabled",type:"boolean"}],D=[{name:"id",type:"string"},{name:"created_at",type:"string"},{name:"updated_at",type:"string"}],o=u.filter(i=>!["is_enabled","id","created_at","updated_at"].includes(i.name));o.length===0&&u.length===0&&o.push({name:"name",type:"string"},{name:"description",type:"string"},{name:"value",type:"number"});const R=[...m,...o,...D],S={values:R.map(i=>({userEnteredValue:{stringValue:i.name}}))},_={values:R.map(i=>{let t="";if(i.name==="id")t="demo_01";else if(i.name==="created_at"||i.name==="updated_at")t=new Date().toISOString();else{switch(i.type){case"number":t=123;break;case"boolean":t=!0;break;case"string":t="demo_content";break;default:t=""}i.name==="name"&&(t="範例項目"),i.name==="description"&&(t="請在第一列↑定義欄位名稱(Key)，從第二列開始輸入您的資料。"),i.name==="status"&&(t="active"),i.name==="is_enabled"&&(t=!0)}return typeof t=="boolean"?{userEnteredValue:{boolValue:t}}:typeof t=="number"?{userEnteredValue:{numberValue:t}}:{userEnteredValue:{stringValue:String(t)}}})},y=[{properties:{title:p,sheetId:0},data:[{startRow:0,startColumn:0,rowData:[S,_]}]}],I=await fetch(O.SHEETS,{method:"POST",headers:U(e),body:JSON.stringify({properties:{title:w},sheets:y})});if(!I.ok){const i=await I.json().catch(()=>({}));throw new Error(i.error?.message||"Failed to create spreadsheet")}const A=await I.json(),b=A.spreadsheetId,v=A.spreadsheetUrl,$=R.map((i,t)=>({...i,index:t})).filter(i=>i.type==="boolean");if($.length>0){const i=$.map(t=>({setDataValidation:{range:{sheetId:0,startRowIndex:1,startColumnIndex:t.index,endColumnIndex:t.index+1,endRowIndex:2},rule:{condition:{type:"BOOLEAN"},strict:!0,showCustomUi:!0}}}));await fetch(`${O.SHEETS}/${b}:batchUpdate`,{method:"POST",headers:U(e),body:JSON.stringify({requests:i})}).catch(t=>console.error("Failed to set validation",t))}return{id:b,spreadsheetUrl:v}}async function Q(e,w,u){const p=await fetch(O.SCRIPT_PROJECTS,{method:"POST",headers:U(e),body:JSON.stringify({title:u,parentId:w})});if(!p.ok){const D=await p.json().catch(()=>({}));throw new Error(D.error?.message||"Failed to create script project")}return(await p.json()).scriptId}async function X(e,w){const u=`${O.SCRIPT_PROJECTS}/${w}/content`,p=K(),m={timeZone:"Asia/Taipei",oauthScopes:["https://www.googleapis.com/auth/spreadsheets.currentonly"],runtimeVersion:"V8",webapp:{access:"ANYONE_ANONYMOUS",executeAs:"USER_DEPLOYING"}},D=await fetch(u,{method:"PUT",headers:U(e),body:JSON.stringify({files:[{name:"appsscript",type:"JSON",source:JSON.stringify(m)},{name:"Code",type:"SERVER_JS",source:p}]})});if(!D.ok){const o=await D.json().catch(()=>({}));throw new Error(o.error?.message||"Failed to update script content")}}async function ee(e,w){const u=`${O.SCRIPT_PROJECTS}/${w}/versions`,p=`${O.SCRIPT_PROJECTS}/${w}/deployments`,m=U(e),D=await fetch(u,{method:"POST",headers:m,body:JSON.stringify({description:"Initial Version"})});if(!D.ok){const y=await D.json().catch(()=>({}));throw new Error(y.error?.message||"Failed to create script version")}const R=(await D.json()).versionNumber,S=await fetch(p,{method:"POST",headers:m,body:JSON.stringify({versionNumber:R,manifestFileName:"appsscript"})});if(!S.ok){const y=await S.json().catch(()=>({}));throw new Error(y.error?.message||"Failed to deploy web app")}return(await S.json()).entryPoints[0].webApp.url}async function te(e,w,u){const p=`${O.DRIVE_FILES}/${w}`;await fetch(p,{method:"PATCH",headers:U(e),body:JSON.stringify({description:u})})}var q={};Y(q,{useGoogleAuth:()=>ae,useSheetManager:()=>ue});var re=["https://www.googleapis.com/auth/script.projects","https://www.googleapis.com/auth/script.deployments","https://www.googleapis.com/auth/drive.file"].join(" "),ae=({clientId:e})=>{const[w,u]=C.useState(null),[p,m]=C.useState(null),[D,o]=C.useState(!1),[R,S]=C.useState(""),[_,y]=C.useState(void 0),[I,A]=C.useState(!1),b=C.useCallback(async t=>{A(!0);try{const x=await fetch(`https://script.googleapis.com/v1/projects?_t=${Date.now()}`,{method:"POST",headers:{Authorization:`Bearer ${t}`,"Content-Type":"application/json"},body:JSON.stringify({title:"Auth Check (Auto Delete)"})});if(x.status===403)console.log("Apps Script API create check failed (403), assuming disabled."),y(!1);else if(x.ok){y(!0);const j=(await x.json()).scriptId;if(j)try{await fetch(`https://www.googleapis.com/drive/v3/files/${j}`,{method:"DELETE",headers:{Authorization:`Bearer ${t}`}})}catch(M){console.warn("Failed to cleanup check file:",M)}}else console.warn("Apps Script check returned unexpected status:",x.status),y(!0)}catch(x){console.error("檢查 Apps Script 狀態失敗:",x),y(!0)}finally{A(!1)}},[]),v=C.useCallback(t=>{if(t.error){S(`授權失敗: ${t.error}`),o(!1);return}m(t.access_token),b(t.access_token),S(""),o(!1)},[b]);C.useEffect(()=>{if(!e)return;const t=document.createElement("script");t.src="https://accounts.google.com/gsi/client",t.async=!0,t.defer=!0,t.onload=()=>{const x=window.google.accounts.oauth2.initTokenClient({client_id:e,scope:re,callback:N=>v(N)});u(x)},document.body.appendChild(t)},[v,e]);const $=()=>{if(!w){S("Google SDK 尚未載入完成");return}o(!0),w.requestAccessToken({prompt:"consent"})},i=C.useCallback(()=>{p&&b(p)},[p,b]);return{accessToken:p,login:$,loading:D,error:R,isAppsScriptEnabled:_,setIsAppsScriptEnabled:y,isChecking:I,recheckAuth:i}},ue=e=>{const[w,u]=C.useState(!1),[p,m]=C.useState("idle"),[D,o]=C.useState(""),[R,S]=C.useState(null),[_,y]=C.useState([]),[I,A]=C.useState(""),[b,v]=C.useState("");return{loading:w,error:D,files:_,creationResult:R,testData:I,authUrl:b,createSheet:async h=>{const{sheetName:f,tabName:a="default",prefix:r="vibesheet-",columns:g=[{name:"name",type:"string"},{name:"value",type:"number"}]}=h;if(e){if(!f.trim()){o("請輸入表格名稱");return}u(!0),m("creating_sheet"),o(""),S(null);try{const c=`${r}${f}`,{id:n,spreadsheetUrl:d}=await Z(e,c,g,a);m("creating_script");const s=await Q(e,n,c);m("updating_script"),await X(e,s),m("deploying");const l=await ee(e,s);m("finishing");const F=JSON.stringify({scriptId:s,scriptUrl:l});await te(e,n,F),S({success:!0,spreadsheetUrl:d,spreadsheetId:n,scriptUrl:l,tip:"請用擁有者 Google 帳號在瀏覽器開啟 scriptUrl 並完成授權，否則匿名存取會被 Google 拒絕 (403)。"}),m("completed")}catch(c){console.error(c),o(c.message||"建立資源時發生未知錯誤"),m("idle")}finally{u(!1)}}},fetchFiles:async(h="vibesheet-")=>{if(e){u(!0),o(""),v(""),A("");try{const f=`name contains '${h}' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`,a=`https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(f)}&fields=files(id, name, webViewLink, description)`,r=await fetch(a,{headers:{Authorization:`Bearer ${e}`}});if(!r.ok){const n=await r.json().catch(()=>({}));throw console.error("Drive API Error:",n),new Error(n.error?.message||`請求失敗 (${r.status}): 請確認 Google Drive API 已啟用`)}const c=((await r.json()).files||[]).map(n=>{let d={},s=!0;if(n.description)try{d=JSON.parse(n.description),d.scriptUrl&&(s=!1)}catch{console.warn("Metadata parse failed:",n.id)}return{...n,...d,isError:s}});y(c)}catch(f){o(f.message||"取得列表失敗")}finally{u(!1)}}},testConnection:async(h,f="")=>{if(!e)return;u(!0),A(""),o(""),v("");let a=h.scriptUrl||"";try{if(!a&&h.description)try{const s=JSON.parse(h.description);s.scriptUrl&&(a=s.scriptUrl)}catch(s){console.error(s)}if(!a){const s=`'${h.id}' in parents and mimeType = 'application/vnd.google-apps.script' and trashed = false`,l=`https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(s)}&fields=files(id,name)`,F=await fetch(l,{headers:{Authorization:`Bearer ${e}`}});if(!F.ok)throw new Error("無法搜尋關聯的 Apps Script");const B=await F.json();if(!B.files||B.files.length===0)throw new Error("找不到關聯的 Script");const J=`https://script.googleapis.com/v1/projects/${B.files[0].id}/deployments`,P=await fetch(J,{headers:{Authorization:`Bearer ${e}`}});if(P.ok){const T=(await P.json()).deployments?.find(V=>V.entryPoints?.some(z=>z.entryPointType==="WEB_APP"));T&&(a=T.entryPoints[0].webApp.url)}}if(!a)throw new Error("無法取得 Script URL");let r=`${a}${a.includes("?")?"&":"?"}t=${new Date().getTime()}`;f&&f.trim()&&(r+=`&fields=${encodeURIComponent(f.trim())}`);const g=await fetch(r,{method:"GET",redirect:"follow",credentials:"omit"});if(!g.ok)throw new Error(`Script 請求失敗 (${g.status})`);const c=g.headers.get("content-type");if(!c||!c.includes("application/json")){const s=await g.text();if(s.trim().startsWith("<"))throw console.error("Script returned HTML:",s),new Error(`連線失敗 (CORS/權限問題)。請確認：
1. 您是否已建立新的表格？(舊表格的 Script 權限未更新)
2. Script 是否部署為「任何人 (含匿名)」？`);try{const l=JSON.parse(s);if(l.error)throw new Error(`Script 回傳錯誤: ${l.error}`);A(JSON.stringify(l.data||l,null,2));return}catch{throw new Error(`回傳格式錯誤: ${s.substring(0,100)}...`)}}const n=await g.json();if(n.error)throw new Error(`Script 回傳錯誤: ${n.error}`);const d=n.data;!d||Array.isArray(d)&&d.length===0?A("[] (目前無資料)"):A(JSON.stringify(d,null,2))}catch(r){console.error(r),r.message==="Failed to fetch"||r.message.includes("CORS")||r.message.includes("HTML")||r.message.includes("403")||r.message.includes("連線失敗")?(o("需要授權：Google 要求您必須手動允許此腳本執行。"),v(a)):o(`測試失敗: ${r.message}`)}finally{u(!1)}},addTestData:async(h,f=1)=>{if(!e)return;u(!0),A(""),o(""),v("");let a=h.scriptUrl||"";try{if(!a&&h.description)try{const d=JSON.parse(h.description);d.scriptUrl&&(a=d.scriptUrl)}catch(d){console.error(d)}if(!a){const d=`'${h.id}' in parents and mimeType = 'application/vnd.google-apps.script' and trashed = false`,s=`https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(d)}&fields=files(id,name)`,l=await fetch(s,{headers:{Authorization:`Bearer ${e}`}});if(!l.ok)throw new Error("無法搜尋關聯的 Apps Script");const F=await l.json();if(!F.files||F.files.length===0)throw new Error("找不到關聯的 Script");const G=`https://script.googleapis.com/v1/projects/${F.files[0].id}/deployments`,J=await fetch(G,{headers:{Authorization:`Bearer ${e}`}});if(J.ok){const k=(await J.json()).deployments?.find(T=>T.entryPoints?.some(V=>V.entryPointType==="WEB_APP"));k&&(a=k.entryPoints[0].webApp.url)}}if(!a)throw new Error("無法取得 Script URL");const r=`${a}${a.includes("?")?"&":"?"}t=${new Date().getTime()}`,g=Array.from({length:f}).map(()=>({name:`Test Item ${Math.floor(Math.random()*1e3)}`,value:Math.floor(Math.random()*100),note:`Added by Vibe Coding on ${new Date().toLocaleTimeString()}`})),c=await fetch(r,{method:"POST",body:JSON.stringify(g)});if(!c.ok)throw new Error(`Script 請求失敗 (${c.status})`);const n=await c.json();if(n.error)throw new Error(`Script 回傳錯誤: ${n.error}`);A(JSON.stringify(n,null,2))}catch(r){console.error(r),r.message==="Failed to fetch"||r.message.includes("CORS")||r.message.includes("HTML")||r.message.includes("403")||r.message.includes("連線失敗")?(o("需要授權 (POST)：Google 要求您必須手動允許此腳本執行。"),v(a)):o(`新增失敗: ${r.message}`)}finally{u(!1)}},updateTestData:async(h,f)=>{if(!e)return;u(!0),A(""),o("");let a=h.scriptUrl||"";try{if(!a&&h.description)try{a=JSON.parse(h.description).scriptUrl||""}catch(F){console.error(F)}if(!a)throw new Error("無法取得 Script URL，請確認表格建立正確。");const r=`${a}${a.includes("?")?"&":"?"}t=${new Date().getTime()}`;let g=f,c={};if(!g){const F=await fetch(r);if(!F.ok)throw new Error("無法讀取現有資料以進行更新");const B=await F.json();if(!B.data||!Array.isArray(B.data)||B.data.length===0)throw new Error("表格目前是空的，請先新增資料再測試更新。");if(c=B.data[0],g=c.id,!g)throw new Error("資料中找不到 id 欄位，無法進行更新")}const n={id:g,name:f?"Updated via ID input":`${c?.name||"Item"} (Updated ${new Date().toLocaleTimeString()})`,value:Math.floor(Math.random()*9999)},d=`${r}&method=PUT`,s=await fetch(d,{method:"POST",body:JSON.stringify(n)});if(!s.ok)throw new Error(`Update 請求失敗 (${s.status})`);const l=await s.json();if(l.error)throw new Error(`Script Update Error: ${l.error}`);A(JSON.stringify({action:"Update Row",targetId:g,sentPayload:n,response:l},null,2))}catch(r){console.error(r),o(`更新失敗: ${r.message}`)}finally{u(!1)}},deleteTestData:async h=>{if(!e)return;u(!0),A(""),o("");let f=h.scriptUrl||"";try{if(!f&&h.description)try{f=JSON.parse(h.description).scriptUrl||""}catch(l){console.error(l)}if(!f)throw new Error("無法取得 Script URL");const a=`${f}${f.includes("?")?"&":"?"}t=${new Date().getTime()}`,r=await fetch(a);if(!r.ok)throw new Error("無法讀取資料");const g=await r.json();if(!g.data||!Array.isArray(g.data)||g.data.length===0)throw new Error("無資料可刪除");const c=g.data.find(l=>l.is_enabled!==!1&&l.is_enabled!=="FALSE");if(!c)throw new Error("找不到有效 (is_enabled=true) 的資料可刪除，或資料皆已刪除。");const n=`${a}&method=DELETE`,d=await fetch(n,{method:"POST",body:JSON.stringify({id:c.id})});if(!d.ok)throw new Error(`Delete 請求失敗 (${d.status})`);const s=await d.json();if(s.error)throw new Error(`Script Delete Error: ${s.error}`);A(JSON.stringify({action:"Soft Delete Row",targetId:c.id,response:s},null,2))}catch(a){console.error(a),o(`刪除失敗: ${a.message}`)}finally{u(!1)}},resetCreation:()=>S(null),clearTestData:()=>A(""),clearError:()=>o(""),clearAuthUrl:()=>v(""),creationStatus:p}};const se=C.lazy(()=>H(()=>import("./SheetSelector-DWJA5MHS.js"),__vite__mapDeps([0,1,2]),import.meta.url)),oe="413406611086-rctqdeav5iirjq1f6p7umd414gfc3uiu.apps.googleusercontent.com",ne=({onScriptSelected:e})=>{const{login:w,accessToken:u,isAppsScriptEnabled:p}=q.useGoogleAuth({clientId:oe});if(u){if(u&&!p)return E.jsx("div",{className:"card",style:{backgroundColor:"var(--bg-card)",color:"var(--text-main)",border:"1px solid var(--border-color)",borderRadius:"8px"},children:p===void 0?E.jsx("p",{children:"Checking your Apps Script permission..."}):E.jsxs(E.Fragment,{children:[E.jsx("p",{children:"Before we start using MoneyBook, you need to enable Apps Script."}),E.jsx("a",{target:"_blank",href:"https://script.google.com/home/usersettings?pli=1",children:"Click Me"})]})})}else return E.jsxs("div",{className:"card",style:{color:"var(--text-main)",textAlign:"center",padding:"40px 20px"},children:[E.jsx("img",{src:"icons/icon-192x192.png",alt:"MoneyBook Logo",style:{width:"100px",height:"100px",borderRadius:"20px",marginBottom:"20px",boxShadow:"0 4px 10px var(--shadow-color)"}}),E.jsx("h1",{style:{fontSize:"1.8em",marginBottom:"30px"},children:"Welcome to MoneyBook"}),E.jsx("button",{onClick:w,style:{fontSize:"1.1em",padding:"12px 24px",backgroundColor:"var(--primary)",color:"var(--bg-card)",border:"none",borderRadius:"8px",cursor:"pointer"},children:"Login with Google"})]});return E.jsx(C.Suspense,{fallback:E.jsx("div",{style:{display:"flex",justifyContent:"center",alignItems:"center",height:"100vh"},children:E.jsx("span",{className:"spinner",style:{width:"30px",height:"30px",border:"3px solid var(--primary)",borderTopColor:"transparent",borderRadius:"50%",animation:"spin 1s linear infinite"}})}),children:E.jsx(se,{token:u,onSelect:m=>{localStorage.setItem("vibe_script_url",m),e(m)}})})},pe=Object.freeze(Object.defineProperty({__proto__:null,default:ne},Symbol.toStringTag,{value:"Module"}));export{pe as i,q as r};
