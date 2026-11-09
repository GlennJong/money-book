const __vite__mapDeps=(i,m=__vite__mapDeps,d=(m.f||(m.f=["./SheetSelector-CIygX112.js","./index-DxyKYn4n.js","./index-D_NmA3Ee.css"])))=>i.map(i=>d[i]);
import{r as y,j as b,_ as Q}from"./index-DxyKYn4n.js";var X=Object.defineProperty,ee=(e,u)=>{for(var r in u)X(e,r,{get:u[r],enumerable:!0})},j={doGet:`
    function doGet(e) {
      try {
        // 授權用連結：擁有者首次開啟 ?action=authorize 完成授權後，顯示提示而非資料
        if ((e.parameter.method || e.parameter.action) === 'authorize') {
          return ContentService.createTextOutput('Authorization successful. You can close this window.');
        }

        // 開發者自訂路由優先 (見 dispatch)
        var routed = vibeDispatch('GET', e);
        if (routed) return routed;

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
           allowedFields = fieldsParam.split(/[,\\s+]+/).filter(function(f) { return f && f.trim().length > 0; });
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
      var crudLock = null;
      try {
        // 開發者自訂路由優先 (見 dispatch)
        var routed = vibeDispatch('POST', e);
        if (routed) return routed;

        // 內建 CRUD 寫入一律排隊執行，避免同時新增時計算出同一個起始列而互相覆蓋
        crudLock = LockService.getScriptLock();
        if (!crudLock.tryLock(10000)) {
          crudLock = null;
          return createJsonResponse({ error: 'LOCK_TIMEOUT' });
        }

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
        var skippedIds = [];

        // 冪等：前端帶入的 id 若已存在 (例如回應遺失後重送)，跳過不重複新增
        var existingIds = {};
        var idColumn = headers.indexOf('id');
        if (idColumn !== -1 && sheet.getMaxRows() > 1) {
            sheet.getRange(2, idColumn + 1, sheet.getMaxRows() - 1, 1).getValues().forEach(function(r) {
                if (r[0] !== '' && r[0] !== null) existingIds[String(r[0])] = true;
            });
        }
        
        incomingRows.forEach(function(rowObj) {
            if (idColumn !== -1 && rowObj.id) {
                var incomingId = String(rowObj.id);
                if (existingIds[incomingId]) {
                    skippedIds.push(incomingId);
                    return;
                }
                existingIds[incomingId] = true;
            }

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
            
            // 超出表格列數時先加列，否則寫入會失敗 (Sheets 預設只有 1000 列)
            var lastRowNeeded = startRow + newRows.length - 1;
            if (lastRowNeeded > sheet.getMaxRows()) {
                sheet.insertRowsAfter(sheet.getMaxRows(), lastRowNeeded - sheet.getMaxRows());
            }

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
           createdIds: createdIds,
           skippedIds: skippedIds
        });
        
      } catch (err) {
        return createJsonResponse({ error: err.toString() });
      } finally {
        if (crudLock) crudLock.releaseLock();
      }
    }
  `,dispatch:`
    var VIBE_RESERVED_ACTIONS = ['PUT', 'UPDATE', 'DELETE', 'authorize'];

    function vibeDispatch(httpMethod, e) {
      var action = e.parameter.method || e.parameter.action;
      if (!action || VIBE_RESERVED_ACTIONS.indexOf(action) !== -1) return null;
      if (typeof VIBE_ROUTES === 'undefined' || !VIBE_ROUTES[httpMethod]) return null;

      var handler = VIBE_ROUTES[httpMethod][action];
      if (typeof handler !== 'function') return null;

      return createJsonResponse(handler(e));
    }
  `,onEdit:`
    function onEdit(e) {
      try {
        vibeFillSystemColumns(e);
      } catch (err) {
        console.error('vibeFillSystemColumns failed: ' + err);
      }
      if (typeof VIBE_ON_EDIT === 'function') VIBE_ON_EDIT(e);
    }

    function vibeFillSystemColumns(e) {
      if (!e || !e.range) return;
      var sheet = e.range.getSheet();
      var lastCol = sheet.getLastColumn();
      if (lastCol === 0) return;

      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var idIndex = headers.indexOf('id');
      var createdIndex = headers.indexOf('created_at');
      var updatedIndex = headers.indexOf('updated_at');
      var enabledIndex = headers.indexOf('is_enabled');
      if (idIndex === -1 && createdIndex === -1 && updatedIndex === -1) return;

      var autoIndexes = [idIndex, createdIndex, updatedIndex];
      var editedOnlyAutoColumns = true;
      for (var c = e.range.getColumn(); c <= e.range.getLastColumn(); c++) {
        if (autoIndexes.indexOf(c - 1) === -1) {
          editedOnlyAutoColumns = false;
          break;
        }
      }
      if (editedOnlyAutoColumns) return;

      var startRow = Math.max(e.range.getRow(), 2);
      var endRow = Math.min(e.range.getLastRow(), sheet.getLastRow());
      if (endRow < startRow) return;

      var numRows = endRow - startRow + 1;
      var rows = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
      var now = new Date().toISOString();
      var systemIndexes = [idIndex, createdIndex, updatedIndex, enabledIndex];

      var touched = false;
      var addedRow = false;
      rows.forEach(function(row) {
        var hasContent = row.some(function(value, i) {
          // 沒勾的勾選框是 FALSE，不算內容
          return systemIndexes.indexOf(i) === -1 && value !== '' && value !== null && value !== false;
        });
        if (!hasContent) return;

        if (idIndex !== -1 && !row[idIndex]) {
          row[idIndex] = Utilities.getUuid();
          if (createdIndex !== -1 && !row[createdIndex]) row[createdIndex] = now;
          if (enabledIndex !== -1) row[enabledIndex] = true;
          addedRow = true;
        }
        if (createdIndex !== -1 && !row[createdIndex]) row[createdIndex] = now;
        if (updatedIndex !== -1) row[updatedIndex] = now;
        touched = true;
      });
      if (!touched) return;

      // 使用者在預留列數之外新增的列沒有勾選框，補上
      if (addedRow && enabledIndex !== -1) {
        var checkbox = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
        sheet.getRange(startRow, enabledIndex + 1, numRows, 1).setDataValidation(checkbox);
      }

      // 只寫回系統欄位，避免覆蓋公式欄或使用者正在編輯的儲存格
      systemIndexes.forEach(function(index) {
        if (index === -1) return;
        var values = rows.map(function(row) { return [row[index]]; });
        sheet.getRange(startRow, index + 1, numRows, 1).setValues(values);
      });
    }
  `,helpers:`
    function createJsonResponse(data) {
      return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
    }
  `},G=e=>`${e}${e.includes("?")?"&":"?"}action=authorize`,te=()=>[j.doGet,j.doPost,j.dispatch,j.onEdit,j.helpers].join(`

`),$={SHEETS:"https://sheets.googleapis.com/v4/spreadsheets",SCRIPT_PROJECTS:"https://script.googleapis.com/v1/projects",DRIVE_FILES:"https://www.googleapis.com/drive/v3/files"},N=e=>({Authorization:`Bearer ${e}`,"Content-Type":"application/json"}),re=[{name:"is_enabled",type:"boolean",note:"勾選 = 啟用"}],ae=[{name:"id",type:"string",readonly:!0,note:"系統自動填入"},{name:"created_at",type:"string",readonly:!0,note:"系統自動填入"},{name:"updated_at",type:"string",readonly:!0,note:"系統自動填入"}],ue=["is_enabled","id","created_at","updated_at"],se=["https://www.googleapis.com/auth/spreadsheets.currentonly"],ne={red:.93,green:.93,blue:.93},oe=e=>{let u="",r=e+1;for(;r>0;){const n=(r-1)%26;u=String.fromCharCode(65+n)+u,r=Math.floor((r-1)/26)}return u},ie=e=>/^\s*ARRAYFORMULA\s*\(/i.test(e)?e:`ARRAYFORMULA(${e})`,de=e=>`'${e.replace(/'/g,"''")}'`,le=e=>typeof e=="boolean"?{boolValue:e}:typeof e=="number"?{numberValue:e}:{stringValue:e==null?"":String(e)},ce=e=>{if(e.systemColumns===!1)return e.columns;const u=e.columns.filter(r=>!ue.includes(r.name));return[...re,...u,...ae]};async function pe(e,u,r,n){if(r.length===0)throw new Error("At least one tab is required");const h=r.map((a,p)=>({...a,sheetId:p,allColumns:ce(a)})),F=(a,p)=>{const l=h.find(o=>o.name===a);if(!l)throw new Error(`Unknown tab "${a}"`);const t=l.allColumns.findIndex(o=>o.name===p);if(t===-1)throw new Error(`Unknown column "${p}" in tab "${a}"`);return t},s=(a,p)=>a.replace(/\{\{\s*([^.{}]*)\.([^{}]+?)\s*\}\}/g,(l,t,o)=>{const v=t.trim()||p;if(!v)throw new Error(`Placeholder {{.${o}}} needs a current tab`);const B=oe(F(v,o.trim()));return`${de(v)}!${B}2:${B}`}),R=h.map(a=>{const p={values:a.allColumns.map(o=>({userEnteredValue:o.arrayFormula?{formulaValue:`={"${o.name}"; ${ie(s(o.arrayFormula,a.name))}}`}:{stringValue:o.name},userEnteredFormat:{textFormat:{bold:!0}},...o.note?{note:o.note}:{}}))},l=new Date().toISOString(),t=(a.demoRows||[]).map(o=>({values:a.allColumns.map(v=>{if(v.arrayFormula)return{};let B=o[v.name];return a.systemColumns!==!1&&(v.name==="id"&&!B&&(B=crypto.randomUUID()),(v.name==="created_at"||v.name==="updated_at")&&!B&&(B=l),v.name==="is_enabled"&&B===void 0&&(B=!0)),{userEnteredValue:le(B)}})}));return{properties:{title:a.name,sheetId:a.sheetId,hidden:!!a.hidden,gridProperties:{frozenRowCount:a.frozenRows??1,columnCount:Math.max(26,a.allColumns.length)}},data:[{startRow:0,startColumn:0,rowData:[p,...t]}]}}),D=await fetch($.SHEETS,{method:"POST",headers:N(e),body:JSON.stringify({properties:{title:u},sheets:R})});if(!D.ok){const a=await D.json().catch(()=>({}));throw new Error(a.error?.message||"Failed to create spreadsheet")}const T=await D.json(),I=T.spreadsheetId,_=T.spreadsheetUrl,f=[];if(h.forEach(a=>{a.allColumns.forEach((p,l)=>{const t={sheetId:a.sheetId,startRowIndex:1,startColumnIndex:l,endColumnIndex:l+1},o=p.validation||(p.type==="boolean"?{kind:"checkbox"}:void 0),v=a.demoRows?.length??0,B=o?.kind==="checkbox"&&v===0;if(o&&!B){let k;o.kind==="checkbox"?k={condition:{type:"BOOLEAN"},strict:!0,showCustomUi:!0}:o.kind==="list"?k={condition:{type:"ONE_OF_LIST",values:o.values.map(q=>({userEnteredValue:q}))},strict:o.strict??!0,showCustomUi:!0}:k={condition:{type:"ONE_OF_RANGE",values:[{userEnteredValue:`=${s(o.range,a.name)}`}]},strict:o.strict??!0,showCustomUi:!0},f.push({setDataValidation:{range:o.kind==="checkbox"?{...t,endRowIndex:1+v}:t,rule:k}})}(p.readonly||p.arrayFormula)&&(f.push({repeatCell:{range:t,cell:{userEnteredFormat:{backgroundColor:ne}},fields:"userEnteredFormat.backgroundColor"}}),f.push({addProtectedRange:{protectedRange:{range:{...t,startRowIndex:0},description:`${p.name} 由系統維護`,warningOnly:!0}}}))})}),n){const a=Object.fromEntries(h.map(p=>[p.name,p.sheetId]));f.push(...n({sheetIds:a,columnIndex:F,resolveRange:s}))}if(f.length>0){const a=await fetch(`${$.SHEETS}/${I}:batchUpdate`,{method:"POST",headers:N(e),body:JSON.stringify({requests:f})});if(!a.ok){const p=await a.json().catch(()=>({}));throw new Error(p.error?.message||"Failed to set up spreadsheet")}}const U=Object.fromEntries(h.map(a=>[a.name,a.sheetId]));return{id:I,spreadsheetUrl:_,sheetIds:U}}async function he(e,u,r=[],n="default"){const h=[{name:"is_enabled",type:"boolean"}],F=[{name:"id",type:"string"},{name:"created_at",type:"string"},{name:"updated_at",type:"string"}],s=r.filter(l=>!["is_enabled","id","created_at","updated_at"].includes(l.name));s.length===0&&r.length===0&&s.push({name:"name",type:"string"},{name:"description",type:"string"},{name:"value",type:"number"});const R=[...h,...s,...F],D={values:R.map(l=>({userEnteredValue:{stringValue:l.name}}))},T={values:R.map(l=>{let t="";if(l.name==="id")t="demo_01";else if(l.name==="created_at"||l.name==="updated_at")t=new Date().toISOString();else{switch(l.type){case"number":t=123;break;case"boolean":t=!0;break;case"string":t="demo_content";break;default:t=""}l.name==="name"&&(t="範例項目"),l.name==="description"&&(t="請在第一列↑定義欄位名稱(Key)，從第二列開始輸入您的資料。"),l.name==="status"&&(t="active"),l.name==="is_enabled"&&(t=!0)}return typeof t=="boolean"?{userEnteredValue:{boolValue:t}}:typeof t=="number"?{userEnteredValue:{numberValue:t}}:{userEnteredValue:{stringValue:String(t)}}})},I=[{properties:{title:n,sheetId:0},data:[{startRow:0,startColumn:0,rowData:[D,T]}]}],_=await fetch($.SHEETS,{method:"POST",headers:N(e),body:JSON.stringify({properties:{title:u},sheets:I})});if(!_.ok){const l=await _.json().catch(()=>({}));throw new Error(l.error?.message||"Failed to create spreadsheet")}const f=await _.json(),U=f.spreadsheetId,a=f.spreadsheetUrl,p=R.map((l,t)=>({...l,index:t})).filter(l=>l.type==="boolean");if(p.length>0){const l=p.map(t=>({setDataValidation:{range:{sheetId:0,startRowIndex:1,startColumnIndex:t.index,endColumnIndex:t.index+1,endRowIndex:2},rule:{condition:{type:"BOOLEAN"},strict:!0,showCustomUi:!0}}}));await fetch(`${$.SHEETS}/${U}:batchUpdate`,{method:"POST",headers:N(e),body:JSON.stringify({requests:l})}).catch(t=>console.error("Failed to set validation",t))}return{id:U,spreadsheetUrl:a}}async function fe(e,u,r){const n=await fetch($.SCRIPT_PROJECTS,{method:"POST",headers:N(e),body:JSON.stringify({title:r,parentId:u})});if(!n.ok){const F=await n.json().catch(()=>({}));throw new Error(F.error?.message||"Failed to create script project")}return(await n.json()).scriptId}async function W(e,u,r){const n=`${$.SCRIPT_PROJECTS}/${u}/content`,h=te(),F={timeZone:"Asia/Taipei",oauthScopes:Array.from(new Set([...se,...r?.oauthScopes||[]])),runtimeVersion:"V8",webapp:{access:"ANYONE_ANONYMOUS",executeAs:"USER_DEPLOYING"}},s=[{name:"appsscript",type:"JSON",source:JSON.stringify(F)},{name:"Code",type:"SERVER_JS",source:h}];r?.code&&s.push({name:"Custom",type:"SERVER_JS",source:r.code});const R=await fetch(n,{method:"PUT",headers:N(e),body:JSON.stringify({files:s})});if(!R.ok){const D=await R.json().catch(()=>({}));throw new Error(D.error?.message||"Failed to update script content")}}async function H(e,u,r){const n=await fetch(`${$.SCRIPT_PROJECTS}/${u}/versions`,{method:"POST",headers:N(e),body:JSON.stringify({description:r})});if(!n.ok){const F=await n.json().catch(()=>({}));throw new Error(F.error?.message||"Failed to create script version")}return(await n.json()).versionNumber}async function we(e,u){const r=`${$.SCRIPT_PROJECTS}/${u}/deployments`,n=N(e),h=await H(e,u,"Initial Version"),F=await fetch(r,{method:"POST",headers:n,body:JSON.stringify({versionNumber:h,manifestFileName:"appsscript"})});if(!F.ok){const R=await F.json().catch(()=>({}));throw new Error(R.error?.message||"Failed to deploy web app")}return(await F.json()).entryPoints[0].webApp.url}async function Y(e,u,r){const n=`${$.DRIVE_FILES}/${u}`;await fetch(n,{method:"PATCH",headers:N(e),body:JSON.stringify({description:r})})}async function me(e,u,r){if(!u.scriptId)throw new Error("File has no scriptId in its metadata");const n=u.scriptId,h=N(e);await W(e,n,r);const F=await H(e,n,`Upgrade ${r?.version||""}`.trim()),s=await fetch(`${$.SCRIPT_PROJECTS}/${n}/deployments`,{headers:h});if(!s.ok){const t=await s.json().catch(()=>({}));throw new Error(t.error?.message||"Failed to list deployments")}const R=await s.json();let D="";try{D=u.description&&JSON.parse(u.description).scriptUrl||""}catch{D=""}const T=t=>t.entryPoints?.find(o=>o.entryPointType==="WEB_APP")?.webApp?.url,I=(R.deployments||[]).filter(t=>t.deploymentConfig?.versionNumber&&T(t)),_=I.find(t=>T(t)===D)||I[0];if(!_)throw new Error("No web app deployment found");const f=await fetch(`${$.SCRIPT_PROJECTS}/${n}/deployments/${_.deploymentId}`,{method:"PUT",headers:h,body:JSON.stringify({deploymentConfig:{scriptId:n,versionNumber:F,manifestFileName:"appsscript",description:r?.version||""}})});if(!f.ok){const t=await f.json().catch(()=>({}));throw new Error(t.error?.message||"Failed to update deployment")}const a=(await f.json()).entryPoints?.find(t=>t.entryPointType==="WEB_APP")?.webApp.url;let p={};try{p=u.description?JSON.parse(u.description):{}}catch{p={}}const l={...p,scriptId:n,scriptUrl:a||p.scriptUrl,scriptVersion:r?.version};return await Y(e,u.id,JSON.stringify(l)),{scriptUrl:l.scriptUrl,versionNumber:F}}var K={};ee(K,{useGoogleAuth:()=>Ce,useSheetManager:()=>Ee});var ge=["https://www.googleapis.com/auth/script.projects","https://www.googleapis.com/auth/script.deployments","https://www.googleapis.com/auth/drive.file"].join(" "),Ce=({clientId:e})=>{const[u,r]=y.useState(null),[n,h]=y.useState(null),[F,s]=y.useState(!1),[R,D]=y.useState(""),[T,I]=y.useState(void 0),[_,f]=y.useState(!1),U=y.useCallback(async t=>{f(!0);try{const o=await fetch(`https://script.googleapis.com/v1/projects?_t=${Date.now()}`,{method:"POST",headers:{Authorization:`Bearer ${t}`,"Content-Type":"application/json"},body:JSON.stringify({title:"Auth Check (Auto Delete)"})});if(o.status===403)console.log("Apps Script API create check failed (403), assuming disabled."),I(!1);else if(o.ok){I(!0);const B=(await o.json()).scriptId;if(B)try{await fetch(`https://www.googleapis.com/drive/v3/files/${B}`,{method:"DELETE",headers:{Authorization:`Bearer ${t}`}})}catch(k){console.warn("Failed to cleanup check file:",k)}}else console.warn("Apps Script check returned unexpected status:",o.status),I(!0)}catch(o){console.error("檢查 Apps Script 狀態失敗:",o),I(!0)}finally{f(!1)}},[]),a=y.useCallback(t=>{if(t.error){D(`授權失敗: ${t.error}`),s(!1);return}h(t.access_token),U(t.access_token),D(""),s(!1)},[U]);y.useEffect(()=>{if(!e)return;const t=document.createElement("script");t.src="https://accounts.google.com/gsi/client",t.async=!0,t.defer=!0,t.onload=()=>{const o=window.google.accounts.oauth2.initTokenClient({client_id:e,scope:ge,callback:v=>a(v)});r(o)},document.body.appendChild(t)},[a,e]);const p=()=>{if(!u){D("Google SDK 尚未載入完成");return}s(!0),u.requestAccessToken({prompt:"consent"})},l=y.useCallback(()=>{n&&U(n)},[n,U]);return{accessToken:n,login:p,loading:F,error:R,isAppsScriptEnabled:T,setIsAppsScriptEnabled:I,isChecking:_,recheckAuth:l}},Ee=e=>{const[u,r]=y.useState(!1),[n,h]=y.useState("idle"),[F,s]=y.useState(""),[R,D]=y.useState(null),[T,I]=y.useState([]),[_,f]=y.useState(""),[U,a]=y.useState("");return{loading:u,error:F,files:T,creationResult:R,testData:_,authUrl:U,createSheet:async g=>{const{sheetName:C,tabName:i="default",prefix:d="vibesheet-",columns:A=[{name:"name",type:"string"},{name:"value",type:"number"}],tabs:S,script:w,setupRequests:E}=g;if(e){if(!C.trim()){s("請輸入表格名稱");return}r(!0),h("creating_sheet"),s(""),D(null);try{const c=`${d}${C}`,{id:m,spreadsheetUrl:x}=S?await pe(e,c,S,E):await he(e,c,A,i);h("creating_script");const O=await fe(e,m,c);h("updating_script"),await W(e,O,w),h("deploying");const P=await we(e,O);h("finishing");const J=JSON.stringify({scriptId:O,scriptUrl:P,...w?.version?{scriptVersion:w.version}:{}});await Y(e,m,J),D({success:!0,spreadsheetUrl:x,spreadsheetId:m,scriptUrl:P,authorizeUrl:G(P),tip:"請用擁有者 Google 帳號在瀏覽器開啟 authorizeUrl 並完成授權，否則匿名存取會被 Google 拒絕 (403)。"}),h("completed")}catch(c){console.error(c),s(c.message||"建立資源時發生未知錯誤"),h("idle")}finally{r(!1)}}},fetchFiles:async(g="vibesheet-")=>{if(e){r(!0),s(""),a(""),f("");try{const C=`name contains '${g}' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false`,i=`https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(C)}&fields=files(id, name, webViewLink, description)`,d=await fetch(i,{headers:{Authorization:`Bearer ${e}`}});if(!d.ok){const w=await d.json().catch(()=>({}));throw console.error("Drive API Error:",w),new Error(w.error?.message||`請求失敗 (${d.status}): 請確認 Google Drive API 已啟用`)}const S=((await d.json()).files||[]).map(w=>{let E={},c=!0;if(w.description)try{E=JSON.parse(w.description),E.scriptUrl&&(c=!1)}catch{console.warn("Metadata parse failed:",w.id)}return{...w,...E,isError:c}});I(S)}catch(C){s(C.message||"取得列表失敗")}finally{r(!1)}}},testConnection:async(g,C="")=>{if(!e)return;r(!0),f(""),s(""),a("");let i=g.scriptUrl||"";try{if(!i&&g.description)try{const c=JSON.parse(g.description);c.scriptUrl&&(i=c.scriptUrl)}catch(c){console.error(c)}if(!i){const c=`'${g.id}' in parents and mimeType = 'application/vnd.google-apps.script' and trashed = false`,m=`https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(c)}&fields=files(id,name)`,x=await fetch(m,{headers:{Authorization:`Bearer ${e}`}});if(!x.ok)throw new Error("無法搜尋關聯的 Apps Script");const O=await x.json();if(!O.files||O.files.length===0)throw new Error("找不到關聯的 Script");const J=`https://script.googleapis.com/v1/projects/${O.files[0].id}/deployments`,V=await fetch(J,{headers:{Authorization:`Bearer ${e}`}});if(V.ok){const L=(await V.json()).deployments?.find(z=>z.entryPoints?.some(Z=>Z.entryPointType==="WEB_APP"));L&&(i=L.entryPoints[0].webApp.url)}}if(!i)throw new Error("無法取得 Script URL");let d=`${i}${i.includes("?")?"&":"?"}t=${new Date().getTime()}`;C&&C.trim()&&(d+=`&fields=${encodeURIComponent(C.trim())}`);const A=await fetch(d,{method:"GET",redirect:"follow",credentials:"omit"});if(!A.ok)throw new Error(`Script 請求失敗 (${A.status})`);const S=A.headers.get("content-type");if(!S||!S.includes("application/json")){const c=await A.text();if(c.trim().startsWith("<"))throw console.error("Script returned HTML:",c),new Error(`連線失敗 (CORS/權限問題)。請確認：
1. 您是否已建立新的表格？(舊表格的 Script 權限未更新)
2. Script 是否部署為「任何人 (含匿名)」？`);try{const m=JSON.parse(c);if(m.error)throw new Error(`Script 回傳錯誤: ${m.error}`);f(JSON.stringify(m.data||m,null,2));return}catch{throw new Error(`回傳格式錯誤: ${c.substring(0,100)}...`)}}const w=await A.json();if(w.error)throw new Error(`Script 回傳錯誤: ${w.error}`);const E=w.data;!E||Array.isArray(E)&&E.length===0?f("[] (目前無資料)"):f(JSON.stringify(E,null,2))}catch(d){console.error(d),d.message==="Failed to fetch"||d.message.includes("CORS")||d.message.includes("HTML")||d.message.includes("403")||d.message.includes("連線失敗")?(s("需要授權：Google 要求您必須手動允許此腳本執行。"),a(G(i))):s(`測試失敗: ${d.message}`)}finally{r(!1)}},addTestData:async(g,C=1)=>{if(!e)return;r(!0),f(""),s(""),a("");let i=g.scriptUrl||"";try{if(!i&&g.description)try{const E=JSON.parse(g.description);E.scriptUrl&&(i=E.scriptUrl)}catch(E){console.error(E)}if(!i){const E=`'${g.id}' in parents and mimeType = 'application/vnd.google-apps.script' and trashed = false`,c=`https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(E)}&fields=files(id,name)`,m=await fetch(c,{headers:{Authorization:`Bearer ${e}`}});if(!m.ok)throw new Error("無法搜尋關聯的 Apps Script");const x=await m.json();if(!x.files||x.files.length===0)throw new Error("找不到關聯的 Script");const P=`https://script.googleapis.com/v1/projects/${x.files[0].id}/deployments`,J=await fetch(P,{headers:{Authorization:`Bearer ${e}`}});if(J.ok){const M=(await J.json()).deployments?.find(L=>L.entryPoints?.some(z=>z.entryPointType==="WEB_APP"));M&&(i=M.entryPoints[0].webApp.url)}}if(!i)throw new Error("無法取得 Script URL");const d=`${i}${i.includes("?")?"&":"?"}t=${new Date().getTime()}`,A=Array.from({length:C}).map(()=>({name:`Test Item ${Math.floor(Math.random()*1e3)}`,value:Math.floor(Math.random()*100),note:`Added by Vibe Coding on ${new Date().toLocaleTimeString()}`})),S=await fetch(d,{method:"POST",body:JSON.stringify(A)});if(!S.ok)throw new Error(`Script 請求失敗 (${S.status})`);const w=await S.json();if(w.error)throw new Error(`Script 回傳錯誤: ${w.error}`);f(JSON.stringify(w,null,2))}catch(d){console.error(d),d.message==="Failed to fetch"||d.message.includes("CORS")||d.message.includes("HTML")||d.message.includes("403")||d.message.includes("連線失敗")?(s("需要授權 (POST)：Google 要求您必須手動允許此腳本執行。"),a(G(i))):s(`新增失敗: ${d.message}`)}finally{r(!1)}},updateTestData:async(g,C)=>{if(!e)return;r(!0),f(""),s("");let i=g.scriptUrl||"";try{if(!i&&g.description)try{i=JSON.parse(g.description).scriptUrl||""}catch(x){console.error(x)}if(!i)throw new Error("無法取得 Script URL，請確認表格建立正確。");const d=`${i}${i.includes("?")?"&":"?"}t=${new Date().getTime()}`;let A=C,S={};if(!A){const x=await fetch(d);if(!x.ok)throw new Error("無法讀取現有資料以進行更新");const O=await x.json();if(!O.data||!Array.isArray(O.data)||O.data.length===0)throw new Error("表格目前是空的，請先新增資料再測試更新。");if(S=O.data[0],A=S.id,!A)throw new Error("資料中找不到 id 欄位，無法進行更新")}const w={id:A,name:C?"Updated via ID input":`${S?.name||"Item"} (Updated ${new Date().toLocaleTimeString()})`,value:Math.floor(Math.random()*9999)},E=`${d}&method=PUT`,c=await fetch(E,{method:"POST",body:JSON.stringify(w)});if(!c.ok)throw new Error(`Update 請求失敗 (${c.status})`);const m=await c.json();if(m.error)throw new Error(`Script Update Error: ${m.error}`);f(JSON.stringify({action:"Update Row",targetId:A,sentPayload:w,response:m},null,2))}catch(d){console.error(d),s(`更新失敗: ${d.message}`)}finally{r(!1)}},deleteTestData:async g=>{if(!e)return;r(!0),f(""),s("");let C=g.scriptUrl||"";try{if(!C&&g.description)try{C=JSON.parse(g.description).scriptUrl||""}catch(m){console.error(m)}if(!C)throw new Error("無法取得 Script URL");const i=`${C}${C.includes("?")?"&":"?"}t=${new Date().getTime()}`,d=await fetch(i);if(!d.ok)throw new Error("無法讀取資料");const A=await d.json();if(!A.data||!Array.isArray(A.data)||A.data.length===0)throw new Error("無資料可刪除");const S=A.data.find(m=>m.is_enabled!==!1&&m.is_enabled!=="FALSE");if(!S)throw new Error("找不到有效 (is_enabled=true) 的資料可刪除，或資料皆已刪除。");const w=`${i}&method=DELETE`,E=await fetch(w,{method:"POST",body:JSON.stringify({id:S.id})});if(!E.ok)throw new Error(`Delete 請求失敗 (${E.status})`);const c=await E.json();if(c.error)throw new Error(`Script Delete Error: ${c.error}`);f(JSON.stringify({action:"Soft Delete Row",targetId:S.id,response:c},null,2))}catch(i){console.error(i),s(`刪除失敗: ${i.message}`)}finally{r(!1)}},upgradeScript:async(g,C)=>{if(e){r(!0),s("");try{return await me(e,g,C)}catch(i){console.error(i),s(i.message||"升級腳本失敗")}finally{r(!1)}}},resetCreation:()=>D(null),clearTestData:()=>f(""),clearError:()=>s(""),clearAuthUrl:()=>a(""),creationStatus:n}};const Fe=y.lazy(()=>Q(()=>import("./SheetSelector-CIygX112.js"),__vite__mapDeps([0,1,2]),import.meta.url)),Ae="413406611086-rctqdeav5iirjq1f6p7umd414gfc3uiu.apps.googleusercontent.com",De=({onScriptSelected:e})=>{const{login:u,accessToken:r,isAppsScriptEnabled:n}=K.useGoogleAuth({clientId:Ae});if(r){if(r&&!n)return b.jsx("div",{className:"card",style:{backgroundColor:"var(--bg-card)",color:"var(--text-main)",border:"1px solid var(--border-color)",borderRadius:"8px"},children:n===void 0?b.jsx("p",{children:"Checking your Apps Script permission..."}):b.jsxs(b.Fragment,{children:[b.jsx("p",{children:"Before we start using MoneyBook, you need to enable Apps Script."}),b.jsx("a",{target:"_blank",href:"https://script.google.com/home/usersettings?pli=1",children:"Click Me"})]})})}else return b.jsxs("div",{className:"card",style:{color:"var(--text-main)",textAlign:"center",padding:"40px 20px"},children:[b.jsx("img",{src:"icons/icon-192x192.png",alt:"MoneyBook Logo",style:{width:"100px",height:"100px",borderRadius:"20px",marginBottom:"20px",boxShadow:"0 4px 10px var(--shadow-color)"}}),b.jsx("h1",{style:{fontSize:"1.8em",marginBottom:"30px"},children:"Welcome to MoneyBook"}),b.jsx("button",{onClick:u,style:{fontSize:"1.1em",padding:"12px 24px",backgroundColor:"var(--primary)",color:"var(--bg-card)",border:"none",borderRadius:"8px",cursor:"pointer"},children:"Login with Google"})]});return b.jsx(y.Suspense,{fallback:b.jsx("div",{style:{display:"flex",justifyContent:"center",alignItems:"center",height:"100vh"},children:b.jsx("span",{className:"spinner",style:{width:"30px",height:"30px",border:"3px solid var(--primary)",borderTopColor:"transparent",borderRadius:"50%",animation:"spin 1s linear infinite"}})}),children:b.jsx(Fe,{token:r,onSelect:h=>{localStorage.setItem("vibe_script_url",h),e(h)}})})},Be=Object.freeze(Object.defineProperty({__proto__:null,default:De},Symbol.toStringTag,{value:"Module"}));export{Be as i,K as r};
