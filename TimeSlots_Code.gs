// ==============================
// Production Coordinator - Google Apps Script (COMPLETE WITH ALL EMAILS)
// ==============================

var SPREADSHEET_ID = '1UOQ3sIqGxi2j0NEHFVxBoBgDv1A9B26V52QK6WLnp1o';

// AI API Configuration
var AI_API_KEY = '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';
var AI_API_TYPE = 'openai'; // Using OpenAI DALL-E

// ------------------------------
// HTTP HANDLERS
// ------------------------------
function doGet(e) {
  try {
    var type = e && e.parameter && e.parameter.type;
    if (type === 'timeslots') return json(getAllTimeSlots());
    if (type === 'submissions') return json(getAllSubmissions());
    return HtmlService.createHtmlOutput('OK');
  } catch (err) {
    return json({ success: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    var p = e.parameter || {};
    var action = p.action;

    if (!action || action === '') {
      var id = p.id, date = p.date, startTime = p.startTime, endTime = p.endTime, createdAt = p.createdAt;
      if (!id || !date || !startTime || !endTime || !createdAt) return text('Error: Missing required parameters');
      var resSlot = writeTimeSlotToSheet(id, normalizeDateString(date), startTime, endTime, createdAt);
      return text(resSlot.success ? 'Success: Time slot saved' : ('Error: ' + resSlot.error));
    }

    if (action === 'delete') {
      if (!p.id) return text('Error: Missing ID');
      var resDel = deleteTimeSlotFromSheet(p.id);
      return text(resDel.success ? 'Success: Time slot deleted' : ('Error: ' + resDel.error));
    }

    if (action === 'submit') {
      var required = ['submissionId','companyName','email','date','timeSlotId','startTime','endTime','website','foodType','submittedAt'];
      for (var i=0;i<required.length;i++) if (!p[required[i]]) return text('Error: Missing ' + required[i]);
      var resSub = writeSubmissionToSheet(
        p.submissionId, p.companyName, p.email, normalizeDateString(p.date),
        p.timeSlotId, p.startTime, p.endTime, p.website, p.foodType, p.submittedAt
      );
      return text(resSub.success ? 'Success: Submission saved' : ('Error: ' + resSub.error));
    }

    if (action === 'approve') {
      if (!p.submissionId) return text('Error: Missing submission ID');
      var resA = approveSubmission(p.submissionId, p.timeSlotId || null);
      return text(resA.success ? 'Success: Submission approved' : ('Error: ' + resA.error));
    }

    if (action === 'reject') {
      if (!p.submissionId) return text('Error: Missing submission ID');
      var fallback = {
        timeSlotId: p.timeSlotId || '',
        date: p.date || '',
        startTime: p.startTime || '',
        endTime: p.endTime || '',
        companyName: p.companyName || '',
        email: p.email || ''
      };
      // Rejection message is optional - just copy and delete
      var resR = rejectSubmission(p.submissionId, p.rejectionMessage || '', fallback);
      return text(resR.success ? 'Success: Submission rejected' : ('Error: ' + resR.error));
    }

    if (action === 'update-approved') {
      if (!p.submissionId) return text('Error: Missing submission ID');
      if (!p.newTimeSlotId) return text('Error: Missing new time slot ID');
      var resU = updateApprovedBooking(
        p.submissionId,
        p.newDate || '',
        p.newStartTime || '',
        p.newEndTime || '',
        p.newTimeSlotId,
        p.oldTimeSlotId || ''
      );
      return text(resU.success ? 'Success: Approved booking updated' : ('Error: ' + resU.error));
    }

    if (action === 'remove-approved-booking') {
      if (!p.submissionId) return text('Error: Missing submission ID');
      var fallback = {
        timeSlotId: p.timeSlotId || '',
        companyName: p.companyName || '',
        email: p.email || '',
        date: p.date || ''
      };
      var resR = removeApprovedBooking(p.submissionId, p.timeSlotId || '', fallback);
      return text(resR.success ? 'Success: Approved booking removed' : ('Error: ' + resR.error));
    }

    return text('Error: Unknown action');
  } catch (err) {
    return text('Error: ' + err.toString());
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function text(s) {
  return ContentService.createTextOutput(s);
}

function normalizeDateString(dateStr) {
  if (!dateStr) return '';
  var s = (dateStr + '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    var y=d.getFullYear(), m=('0'+(d.getMonth()+1)).slice(-2), dd=('0'+d.getDate()).slice(-2);
    return y+'-'+m+'-'+dd;
  }
  return s;
}

function writeTimeSlotToSheet(id, date, startTime, endTime, createdAt) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName('TimeSlots');
    if (!sh) return {success:false,error:'TimeSlots sheet not found'};
    if (!sh.getRange(1,1,1,6).getValues()[0][0]) {
      sh.getRange(1,1,1,6).setValues([['ID','Date','Start Time','End Time','Created At','Submission ID']]);
    }
    sh.appendRow([id, date, startTime, endTime, createdAt, '']);
    return {success:true};
  } catch (e) { return {success:false,error:e.toString()}; }
}

function deleteTimeSlotFromSheet(id) {
  try {
    var ss=SpreadsheetApp.openById(SPREADSHEET_ID), sh=ss.getSheetByName('TimeSlots');
    if (!sh) return {success:false,error:'TimeSlots sheet not found'};
    var lr=sh.getLastRow(); 
    if (lr<2) return {success:false,error:'No data to delete'};
    
    var idStr = (id + '').trim();
    Logger.log('🔍 DELETE TIME SLOT: Looking for ID: ' + idStr);
    
    // Read Column A (ID column) only
    var data=sh.getRange(2,1,lr-1,1).getValues(); // Column A only (index 0)
    
    for (var i=0;i<data.length;i++) {
      var rowId = (data[i][0] + '').trim();
      
      // Try exact match
      if (rowId === idStr) {
        Logger.log('✅ DELETE TIME SLOT: Found time slot at row ' + (i + 2) + ' by ID match');
        sh.deleteRow(i+2);
        Logger.log('✅ DELETE TIME SLOT: Row deleted from TimeSlots sheet');
        return {success:true};
      }
      
      // Also try as numbers (in case one is number and one is string)
      if (!isNaN(rowId) && !isNaN(idStr)) {
        if (Number(rowId) === Number(idStr)) {
          Logger.log('✅ DELETE TIME SLOT: Found time slot at row ' + (i + 2) + ' by numeric ID match');
          sh.deleteRow(i+2);
          Logger.log('✅ DELETE TIME SLOT: Row deleted from TimeSlots sheet');
          return {success:true};
        }
      }
    }
    
    Logger.log('❌ DELETE TIME SLOT: Time slot not found. ID: ' + idStr);
    return {success:false,error:'Time slot not found. Searched by ID: ' + idStr};
  } catch(e){
    Logger.log('❌❌❌ DELETE TIME SLOT ERROR: ' + e.toString());
    return {success:false,error:e.toString()};
  }
}

function getAllTimeSlots() {
  try {
    var ss=SpreadsheetApp.openById(SPREADSHEET_ID), sh=ss.getSheetByName('TimeSlots');
    if (!sh) return {success:false,error:'TimeSlots sheet not found',timeSlots:[]};
    var lr=sh.getLastRow(); if (lr<2) return {success:true,timeSlots:[]};
    var data=sh.getRange(2,1,lr-1,6).getValues();
    var rows=data.map(function(r){
      var dateVal=r[1], dateStr='';
      if (dateVal instanceof Date) {
        var y=dateVal.getFullYear(), m=('0'+(dateVal.getMonth()+1)).slice(-2), d=('0'+dateVal.getDate()).slice(-2);
        dateStr=y+'-'+m+'-'+d;
      } else if (dateVal) dateStr=normalizeDateString(dateVal);
      return { id:r[0], date:dateStr, startTime:r[2], endTime:r[3], createdAt:r[4], submissionId:r[5]||'' };
    });
    return {success:true,timeSlots:rows};
  } catch(e){return {success:false,error:e.toString(),timeSlots:[]};}
}

function writeSubmissionToSheet(submissionId, companyName, email, date, timeSlotId, startTime, endTime, website, foodType, submittedAt) {
  try {
    var ss=SpreadsheetApp.openById(SPREADSHEET_ID), sh=ss.getSheetByName('Submissions');
    if (!sh) return {success:false,error:'Submissions sheet not found'};
    
    // Check if headers exist, if not create them (14 columns now: Time Slot ID in Column M, Rejection Message in Column N)
    if (!sh.getRange(1,1,1,14).getValues()[0][0]) {
      sh.getRange(1,1,1,14).setValues([[
        'Submission ID','Company Name','Email','Date','Time Slot ID',
        'Start Time','End Time','Website','Food Type','Submitted At',
        'Approved At','Status','Time Slot ID','Rejection Message'
      ]]);
    }
    
    // Write row: timeSlotId goes in Column E (index 4) AND Column M (index 12)
    // Rejection Message in Column N (index 13)
    sh.appendRow([submissionId,companyName,email,date,timeSlotId,startTime,endTime,website,foodType,submittedAt,'','PENDING',timeSlotId,'']);
    return {success:true};
  } catch(e){return {success:false,error:e.toString()};}
}

function getAllSubmissions() {
  try {
    var ss=SpreadsheetApp.openById(SPREADSHEET_ID), sh=ss.getSheetByName('Submissions');
    if (!sh) return {success:false,error:'Submissions sheet not found',submissions:[]};
    
    var lr=sh.getLastRow();
    if (lr<2) {
      return {success:true,submissions:[]};
    }
    
    // Read all rows from row 2 to last row - check if it's 13 or 14 columns
    var numCols = sh.getLastColumn();
    var colCount = numCols >= 14 ? 14 : 13; // Support both old (13) and new (14) column format
    
    var data=sh.getRange(2,1,lr-1,colCount).getValues();
    
    // Map each row to a submission object
    var rows=data.map(function(r, index){
      // Only skip completely empty rows
      var isEmpty = true;
      for (var i=0; i<colCount; i++) {
        if (r[i] !== null && r[i] !== '' && r[i] !== undefined) {
          isEmpty = false;
          break;
        }
      }
      
      if (isEmpty) return null;
      
      var status=r[11]||'PENDING';
      var dateVal=r[3], dateStr='';
      if (dateVal instanceof Date) {
        var y=dateVal.getFullYear(), m=('0'+(dateVal.getMonth()+1)).slice(-2), dd=('0'+dateVal.getDate()).slice(-2);
        dateStr=y+'-'+m+'-'+dd;
      } else if (dateVal) dateStr=normalizeDateString(dateVal);
      
      // Handle both 13 and 14 column formats
      var timeSlotIdM = colCount >= 14 ? (r[12]||'') : ''; // Column M (index 12) - new location
      var rejectionMessage = colCount >= 14 ? (r[13]||'') : (r[12]||''); // Column N (index 13) if 14 cols, else Column M (index 12)
      
      return {
        id:r[0]||('submission_' + (index+2)),
        companyName:r[1]||'',
        email:r[2]||'',
        date:dateStr,
        timeSlotId:r[4]||'', // Column E (index 4) - original location
        timeSlotIdM:timeSlotIdM, // Column M (index 12) - new location
        startTime:r[5]||'',
        endTime:r[6]||'',
        website:r[7]||'',
        foodType:r[8]||'',
        submittedAt:r[9]||'',
        approvedAt:r[10]||'',
        status:status,
        approved:status==='APPROVED',
        rejected:status==='REJECTED',
        rejectionMessage:rejectionMessage
      };
    });
    
    // Filter out null entries
    rows = rows.filter(function(r) { return r !== null; });
    
    return {success:true,submissions:rows};
  } catch(e){
    return {success:false,error:e.toString(),submissions:[]};
  }
}

function approveSubmission(submissionId, timeSlotId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var subSh = ss.getSheetByName('Submissions');
    if (!subSh) return {success:false,error:'Submissions sheet not found'};
    
    var lr = subSh.getLastRow();
    if (lr < 2) return {success:false,error:'No submissions found'};
    
    // Check column count (13 or 14)
    var numCols = subSh.getLastColumn();
    var colCount = numCols >= 14 ? 14 : 13;
    
    // Read ALL data including all columns
    var data = subSh.getRange(2,1,lr-1,colCount).getValues();
    var rowIndex = -1;
    var rowData = null;
    var submissionIdStr = (submissionId + '').trim();
    
    // Log for debugging
    Logger.log('🔍 APPROVAL: Looking for submission ID: ' + submissionIdStr);
    Logger.log('🔍 APPROVAL: Total rows in Submissions: ' + (lr - 1));
    
    // Find the submission by ID
    for (var i = 0; i < data.length; i++) {
      var rowId = (data[i][0] + '').trim();
      if (rowId === submissionIdStr) {
        rowIndex = i + 2; // +2 because row 1 is header, and array is 0-indexed
        rowData = data[i];
        Logger.log('✅ APPROVAL: Found submission at row ' + rowIndex + ' by ID match');
        break;
      }
    }
    
    // If not found, return error
    if (rowIndex === -1 || !rowData) {
      Logger.log('❌ APPROVAL: Submission not found. ID: ' + submissionIdStr);
      return {success:false,error:'Submission not found. Searched by ID: ' + submissionIdStr};
    }
    
    Logger.log('✅ APPROVAL: Found submission at row ' + rowIndex);
    
    // Update Status (column L, index 11) to 'APPROVED'
    subSh.getRange(rowIndex, 12).setValue('APPROVED'); // Column L (index 11 = column 12)
    
    // Set Approved At in Column K (index 10) if not already set
    if (!rowData[10]) {
      subSh.getRange(rowIndex, 11).setValue(new Date().toISOString()); // Column K (index 10 = column 11)
    }
    
    // Update Time Slot with submission ID if timeSlotId provided
    if (timeSlotId) {
      var tsSh = ss.getSheetByName('TimeSlots');
      if (tsSh) {
        if (!tsSh.getRange(1,6).getValue()) tsSh.getRange(1,6).setValue('Submission ID');
        var tlr = tsSh.getLastRow();
        if (tlr >= 2) {
          var tData = tsSh.getRange(2,1,tlr-1,6).getValues();
          var tsRow = -1;
          for (var j=0;j<tData.length;j++) {
            if (((tData[j][0]+'').trim()) === ((timeSlotId+'').trim())) { tsRow = j+2; break; }
          }
          if (tsRow !== -1) {
            tsSh.getRange(tsRow, 6).setValue(submissionId);
            Logger.log('✅ APPROVAL: Updated Time Slot with submission ID');
          }
        }
      }
    }
    
    // Send approval emails
    var emailResult = sendApprovalEmails(submissionId);
    if (!emailResult.success) {
      Logger.log('⚠️ APPROVAL: Email sending failed: ' + emailResult.error);
      // Don't fail the approval if email fails - just log it
    }
    
    Logger.log('✅✅✅ APPROVAL COMPLETE: Success!');
    return {success:true};
  } catch (e) {
    Logger.log('❌❌❌ APPROVAL ERROR: ' + e.toString());
    return {success:false,error:e.toString()};
  }
}

function rejectSubmission(submissionId, rejectionMessage, fallback) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var subSh = ss.getSheetByName('Submissions');
    if (!subSh) return {success:false,error:'Submissions sheet not found'};
    
    var lr = subSh.getLastRow();
    if (lr < 2) return {success:false,error:'No submissions found'};
    
    // Check column count (13 or 14)
    var numCols = subSh.getLastColumn();
    var colCount = numCols >= 14 ? 14 : 13;
    
    // Read ALL data including all columns
    var data = subSh.getRange(2,1,lr-1,colCount).getValues();
    var rowIndex = -1;
    var rowData = null;
    var submissionIdStr = (submissionId + '').trim();
    
    // Log for debugging
    Logger.log('🔍 REJECTION: Looking for submission ID: ' + submissionIdStr);
    Logger.log('🔍 REJECTION: Total rows in Submissions: ' + (lr - 1));
    Logger.log('🔍 REJECTION: Column count: ' + colCount);
    
    // Step 1: Try to find by exact submission ID match
    for (var i = 0; i < data.length; i++) {
      var rowId = (data[i][0] + '').trim();
      if (rowId === submissionIdStr) {
        rowIndex = i + 2; // +2 because row 1 is header, and array is 0-indexed
        rowData = data[i];
        Logger.log('✅ REJECTION: Found submission at row ' + rowIndex + ' by ID match');
        break;
      }
    }
    
    // Step 2: If not found, try fallback matching by timeSlotId
    if (rowIndex === -1 && fallback && fallback.timeSlotId) {
      var timeSlotIdStr = (fallback.timeSlotId + '').trim();
      if (timeSlotIdStr) {
        for (var j = 0; j < data.length; j++) {
          var rowTimeSlotId = (data[j][4] + '').trim(); // Column E (index 4) is Time Slot ID
          if (rowTimeSlotId === timeSlotIdStr) {
            rowIndex = j + 2;
            rowData = data[j];
            Logger.log('✅ REJECTION: Found submission at row ' + rowIndex + ' by timeSlotId match');
            break;
          }
        }
      }
    }
    
    // Step 3: If still not found, try composite matching (date + time + company/email)
    if (rowIndex === -1 && fallback) {
      var fDate = (fallback.date + '').trim();
      var fStartTime = (fallback.startTime + '').trim();
      var fEndTime = (fallback.endTime + '').trim();
      var fCompany = ((fallback.companyName || '') + '').trim().toLowerCase();
      var fEmail = ((fallback.email || '') + '').trim().toLowerCase();
      
      for (var k = 0; k < data.length; k++) {
        var r = data[k];
        var rDate = (r[3] + '').trim(); // Column D (index 3) is Date
        var rStartTime = (r[5] + '').trim(); // Column F (index 5) is Start Time
        var rEndTime = (r[6] + '').trim(); // Column G (index 6) is End Time
        var rCompany = ((r[1] || '') + '').trim().toLowerCase(); // Column B (index 1) is Company Name
        var rEmail = ((r[2] || '') + '').trim().toLowerCase(); // Column C (index 2) is Email
        
        var dateMatch = !fDate || rDate === fDate || normalizeDateString(rDate) === normalizeDateString(fDate);
        var timeMatch = (!fStartTime || rStartTime === fStartTime) && (!fEndTime || rEndTime === fEndTime);
        var whoMatch = !fCompany || rCompany === fCompany || (!fEmail || rEmail === fEmail);
        
        if (dateMatch && timeMatch && whoMatch) {
          rowIndex = k + 2;
          rowData = r;
          Logger.log('✅ REJECTION: Found submission at row ' + rowIndex + ' by composite match');
          break;
        }
      }
    }
    
    // If still not found, return error
    if (rowIndex === -1 || !rowData) {
      Logger.log('❌ REJECTION: Submission not found. ID: ' + submissionIdStr);
      return {success:false,error:'Submission not found. Searched by ID: ' + submissionIdStr + ', timeSlotId: ' + (fallback && fallback.timeSlotId ? fallback.timeSlotId : 'N/A')};
    }
    
    Logger.log('✅ REJECTION: Found submission at row ' + rowIndex);
    
    // Step 4: Copy ALL columns from Submissions to Rejections
    // Always create 14 columns for Rejections sheet (even if Submissions has 13)
    var submissionData = [];
    
    // Copy all columns from Submissions row
    for (var col = 0; col < colCount; col++) {
      if (col < rowData.length) {
        submissionData.push(rowData[col] || '');
      } else {
        submissionData.push('');
      }
    }
    
    // CRITICAL FIX: Pad to 14 columns if needed (for Rejections sheet)
    while (submissionData.length < 14) {
      submissionData.push('');
    }
    
    // Override Status (column L, index 11) to 'REJECTED'
    submissionData[11] = 'REJECTED';
    
    // Set Rejection Message in Column N (index 13) - can be empty if no message provided
    submissionData[13] = (rejectionMessage || '').toString().trim();
    
    // If Time Slot ID is in Column E (index 4), also copy to Column M (index 12) if empty
    if (!submissionData[12] && submissionData[4]) {
      submissionData[12] = submissionData[4]; // Copy Time Slot ID to Column M
    }
    
    Logger.log('📋 REJECTION: Prepared submissionData with ' + submissionData.length + ' columns');
    Logger.log('📋 REJECTION: Status = ' + submissionData[11]);
    
    // Step 5: Get or create Rejections sheet
    var rejSh = ss.getSheetByName('Rejections');
    if (!rejSh) {
      rejSh = ss.insertSheet('Rejections');
      // Always use 14 columns for Rejections sheet
      rejSh.getRange(1,1,1,14).setValues([[
        'Submission ID','Company Name','Email','Date','Time Slot ID',
        'Start Time','End Time','Website','Food Type','Submitted At',
        'Approved At','Status','Time Slot ID','Rejection Message'
      ]]);
      Logger.log('✅ REJECTION: Created Rejections sheet with 14 columns');
    } else {
      // Ensure headers exist - always use 14 columns for Rejections
      var rejColCount = rejSh.getLastColumn();
      if (rejColCount < 14) {
        // Update to 14 columns
        rejSh.getRange(1,1,1,14).setValues([[
          'Submission ID','Company Name','Email','Date','Time Slot ID',
          'Start Time','End Time','Website','Food Type','Submitted At',
          'Approved At','Status','Time Slot ID','Rejection Message'
        ]]);
        Logger.log('✅ REJECTION: Updated Rejections sheet headers to 14 columns');
      }
    }
    
    // Step 6: Copy/paste: Append the complete row to Rejections sheet
    // submissionData should already be 14 columns at this point
    Logger.log('📋 REJECTION: Copying ' + submissionData.length + ' columns to Rejections sheet');
    rejSh.appendRow(submissionData);
    Logger.log('✅ REJECTION: Row copied to Rejections sheet');
    
    // Step 7: Delete the original row from Submissions sheet
    Logger.log('🗑️ REJECTION: Deleting row ' + rowIndex + ' from Submissions sheet');
    subSh.deleteRow(rowIndex);
    Logger.log('✅ REJECTION: Row deleted from Submissions sheet');
    
    // Send rejection email
    var emailResult = sendRejectionEmail(submissionId);
    if (!emailResult.success) {
      Logger.log('⚠️ REJECTION: Email sending failed: ' + emailResult.error);
      // Don't fail the rejection if email fails - just log it
    }
    
    Logger.log('✅✅✅ REJECTION COMPLETE: Success!');
    return {success:true};
  } catch (e) {
    Logger.log('❌❌❌ REJECTION ERROR: ' + e.toString());
    return {success:false,error:e.toString()};
  }
}

function updateApprovedBooking(submissionId, newDate, newStartTime, newEndTime, newTimeSlotId, oldTimeSlotId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var subSh = ss.getSheetByName('Submissions');
    if (!subSh) return {success:false,error:'Submissions sheet not found'};
    
    var lr = subSh.getLastRow();
    if (lr < 2) return {success:false,error:'No submissions found'};
    
    // Check column count (13 or 14)
    var numCols = subSh.getLastColumn();
    var colCount = numCols >= 14 ? 14 : 13;
    
    // Read ALL data including all columns
    var data = subSh.getRange(2,1,lr-1,colCount).getValues();
    var rowIndex = -1;
    var submissionIdStr = (submissionId + '').trim();
    
    Logger.log('🔍 UPDATE APPROVED: Looking for submission ID: ' + submissionIdStr);
    
    // Find the approved booking by submission ID in Submissions sheet
    for (var i = 0; i < data.length; i++) {
      var rowId = (data[i][0] + '').trim();
      if (rowId === submissionIdStr) {
        rowIndex = i + 2; // +2 because row 1 is header, and array is 0-indexed
        Logger.log('✅ UPDATE APPROVED: Found booking at row ' + rowIndex);
        break;
      }
    }
    
    if (rowIndex === -1) {
      Logger.log('❌ UPDATE APPROVED: Booking not found. ID: ' + submissionIdStr);
      return {success:false,error:'Approved booking not found. Searched by ID: ' + submissionIdStr};
    }
    
    Logger.log('✅ UPDATE APPROVED: Found booking at row ' + rowIndex);
    
    // Update Submissions sheet columns:
    // Column D (index 3) = Date
    // Column F (index 5) = Start Time
    // Column G (index 6) = End Time
    // Column M (index 12) = Time Slot ID
    
    if (newDate) {
      subSh.getRange(rowIndex, 4).setValue(normalizeDateString(newDate)); // Column D
      Logger.log('✅ UPDATE APPROVED: Updated Date to ' + normalizeDateString(newDate));
    }
    
    if (newStartTime) {
      subSh.getRange(rowIndex, 6).setValue(newStartTime); // Column F
      Logger.log('✅ UPDATE APPROVED: Updated Start Time to ' + newStartTime);
    }
    
    if (newEndTime) {
      subSh.getRange(rowIndex, 7).setValue(newEndTime); // Column G
      Logger.log('✅ UPDATE APPROVED: Updated End Time to ' + newEndTime);
    }
    
    // Update Column M (Time Slot ID) - index 12
    subSh.getRange(rowIndex, 13).setValue(newTimeSlotId); // Column M (index 12 = column 13)
    Logger.log('✅ UPDATE APPROVED: Updated Time Slot ID (Column M) to ' + newTimeSlotId);
    
    // Also update Column E (Time Slot ID) - index 4, for consistency
    subSh.getRange(rowIndex, 5).setValue(newTimeSlotId); // Column E (index 4 = column 5)
    Logger.log('✅ UPDATE APPROVED: Updated Time Slot ID (Column E) to ' + newTimeSlotId);
    
    // Update TimeSlots sheet:
    // 1. Clear old time slot's Submission ID (Column F)
    // 2. Set new time slot's Submission ID (Column F)
    
    var tsSh = ss.getSheetByName('TimeSlots');
    if (tsSh) {
      if (!tsSh.getRange(1,6).getValue()) tsSh.getRange(1,6).setValue('Submission ID');
      var tlr = tsSh.getLastRow();
      
      if (tlr >= 2) {
        var tData = tsSh.getRange(2,1,tlr-1,6).getValues();
        
        // Clear old time slot's Submission ID if oldTimeSlotId provided
        if (oldTimeSlotId) {
          var oldTimeSlotIdStr = (oldTimeSlotId + '').trim();
          for (var j = 0; j < tData.length; j++) {
            if (((tData[j][0] + '').trim()) === oldTimeSlotIdStr) {
              tsSh.getRange(j + 2, 6).setValue(''); // Clear Column F
              Logger.log('✅ UPDATE APPROVED: Cleared old time slot ' + oldTimeSlotIdStr + ' Submission ID');
              break;
            }
          }
        }
        
        // Set new time slot's Submission ID
        var newTimeSlotIdStr = (newTimeSlotId + '').trim();
        for (var k = 0; k < tData.length; k++) {
          if (((tData[k][0] + '').trim()) === newTimeSlotIdStr) {
            tsSh.getRange(k + 2, 6).setValue(submissionId); // Set Column F
            Logger.log('✅ UPDATE APPROVED: Set new time slot ' + newTimeSlotIdStr + ' Submission ID to ' + submissionId);
            break;
          }
        }
      }
    }
    
    // Send time change email
    var emailResult = sendTimeChangeEmail(submissionId);
    if (!emailResult.success) {
      Logger.log('⚠️ UPDATE APPROVED: Email sending failed: ' + emailResult.error);
      // Don't fail the update if email fails - just log it
    }
    
    Logger.log('✅✅✅ UPDATE APPROVED COMPLETE: Success!');
    return {success:true};
  } catch (e) {
    Logger.log('❌❌❌ UPDATE APPROVED ERROR: ' + e.toString());
    return {success:false,error:e.toString()};
  }
}

function removeApprovedBooking(submissionId, timeSlotId, fallback) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var subSh = ss.getSheetByName('Submissions');
    if (!subSh) return {success:false,error:'Submissions sheet not found'};
    
    var lr = subSh.getLastRow();
    if (lr < 2) return {success:false,error:'No submissions found'};
    
    var submissionIdStr = (submissionId + '').trim();
    Logger.log('🔍 REMOVE APPROVED: Looking for submission ID: ' + submissionIdStr);
    
    // Check column count (13 or 14)
    var numCols = subSh.getLastColumn();
    var colCount = numCols >= 14 ? 14 : 13;
    
    // Step 1: Find submission by ID and GET THE DATA BEFORE DELETING
    var data = subSh.getRange(2, 1, lr - 1, colCount).getValues(); // Get ALL columns
    var rowIndex = -1;
    var submissionData = null;
    
    for (var i = 0; i < data.length; i++) {
      var rowId = (data[i][0] + '').trim();
      
      // Try exact match
      if (rowId === submissionIdStr) {
        rowIndex = i + 2;
        submissionData = {
          id: data[i][0],
          companyName: data[i][1] || '',
          email: data[i][2] || ''
        };
        Logger.log('✅ REMOVE APPROVED: Found submission at row ' + rowIndex + ' by ID match');
        break;
      }
      
      // Also try as numbers (in case one is number and one is string)
      if (!isNaN(rowId) && !isNaN(submissionIdStr)) {
        if (Number(rowId) === Number(submissionIdStr)) {
          rowIndex = i + 2;
          submissionData = {
            id: data[i][0],
            companyName: data[i][1] || '',
            email: data[i][2] || ''
          };
          Logger.log('✅ REMOVE APPROVED: Found submission at row ' + rowIndex + ' by numeric ID match');
          break;
        }
      }
    }
    
    if (rowIndex === -1 || !submissionData) {
      Logger.log('❌ REMOVE APPROVED: Submission not found in Column A. ID: ' + submissionIdStr);
      return {success:false,error:'Submission not found. Searched by ID: ' + submissionIdStr};
    }
    
    // Step 2: Send rejection email BEFORE deleting (so we have the submission data)
    var emailResult = sendRejectionEmailWithData(submissionData.id, submissionData.companyName, submissionData.email);
    if (!emailResult.success) {
      Logger.log('⚠️ REMOVE APPROVED: Email sending failed: ' + emailResult.error);
      // Don't fail the removal if email fails - just log it
    }
    
    // Step 3: Delete the row from Submissions sheet
    Logger.log('🗑️ REMOVE APPROVED: Deleting row ' + rowIndex + ' from Submissions sheet');
    subSh.deleteRow(rowIndex);
    Logger.log('✅ REMOVE APPROVED: Row deleted from Submissions sheet');
    
    // Step 4: Find and clear the submission ID in Column F (Column 6) of TimeSlots sheet
    var tsSh = ss.getSheetByName('TimeSlots');
    if (tsSh) {
      if (!tsSh.getRange(1,6).getValue()) tsSh.getRange(1,6).setValue('Submission ID');
      var tlr = tsSh.getLastRow();
      if (tlr >= 2) {
        var tData = tsSh.getRange(2, 1, tlr - 1, 6).getValues(); // Columns A through F
        var foundInTimeSlots = false;
        
        for (var j = 0; j < tData.length; j++) {
          var colFValue = (tData[j][5] || '').toString().trim(); // Column F (index 5)
          
          // Check if Column F matches the submission ID
          if (colFValue === submissionIdStr) {
            // Clear Column F (Submission ID)
            tsSh.getRange(j + 2, 6).setValue('');
            Logger.log('✅ REMOVE APPROVED: Cleared Column F (Submission ID) at row ' + (j + 2) + ' in TimeSlots sheet');
            foundInTimeSlots = true;
            break;
          }
          
          // Also try numeric comparison
          if (!isNaN(colFValue) && !isNaN(submissionIdStr)) {
            if (Number(colFValue) === Number(submissionIdStr)) {
              tsSh.getRange(j + 2, 6).setValue('');
              Logger.log('✅ REMOVE APPROVED: Cleared Column F (Submission ID) at row ' + (j + 2) + ' in TimeSlots sheet (numeric match)');
              foundInTimeSlots = true;
              break;
            }
          }
        }
        
        if (!foundInTimeSlots) {
          Logger.log('⚠️ REMOVE APPROVED: Submission ID ' + submissionIdStr + ' not found in Column F of TimeSlots sheet');
        }
      }
    }
    
    Logger.log('✅✅✅ REMOVE APPROVED COMPLETE: Success!');
    return {success:true};
  } catch (e) {
    Logger.log('❌❌❌ REMOVE APPROVED ERROR: ' + e.toString());
    return {success:false,error:e.toString()};
  }
}

// Helper function to send rejection email with submission data (for use when submission is deleted)
function sendRejectionEmailWithData(submissionId, companyName, email) {
  try {
    if (!email) {
      Logger.log('⚠️ REJECTION EMAIL: No email provided');
      return {success:false,error:'No email provided'};
    }
    
    Logger.log('📧 REJECTION EMAIL: Sending to ' + companyName + ' at ' + email);
    
    // Get or create "Approval Email" sheet
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var emailSh = ss.getSheetByName('Approval Email');
    if (!emailSh) {
      emailSh = ss.insertSheet('Approval Email');
      emailSh.getRange(1,1,1,4).setValues([['Email', '', '', 'Email Confirmation']]);
      Logger.log('✅ REJECTION EMAIL: Created Approval Email sheet');
    }
    
    // Read CC recipients from Column D (starting row 2)
    var emailCC = [];
    var emailLastRow = emailSh.getLastRow();
    if (emailLastRow >= 2) {
      var emailCCData = emailSh.getRange(2, 4, emailLastRow - 1, 1).getValues();
      for (var k = 0; k < emailCCData.length; k++) {
        var ccEmail = (emailCCData[k][0] || '').toString().trim();
        if (ccEmail && ccEmail.indexOf('@') !== -1) {
          emailCC.push(ccEmail);
        }
      }
    }
    
    Logger.log('📧 REJECTION EMAIL: Found ' + emailCC.length + ' CC recipients');
    
    // Send rejection email
    var emailSubject = 'Screening Time Rejection Email for ' + companyName;
    var emailBody = 'Hi ' + companyName + ',\n\n' +
                    'Tim\'s production team has reviewed your submission and it has been rejected for the time you selected.\n\n' +
                    'I am adding the Ogilvy team here so you can coordinate directly on next steps and decide how to move forward.\n\n' +
                    'Thank you,\nIsaac';
    
    try {
      var emailOptions = {
        to: email,
        subject: emailSubject,
        body: emailBody
      };
      
      if (emailCC.length > 0) {
        emailOptions.cc = emailCC.join(',');
      }
      
      MailApp.sendEmail(emailOptions);
      Logger.log('✅ REJECTION EMAIL: Sent to ' + email + ' with ' + emailCC.length + ' CC recipients');
      return {success:true};
    } catch (e) {
      Logger.log('❌ REJECTION EMAIL ERROR: ' + e.toString());
      return {success:false,error:e.toString()};
    }
  } catch (e) {
    Logger.log('❌❌❌ REJECTION EMAIL ERROR: ' + e.toString());
    return {success:false,error:e.toString()};
  }
}

// Helper function to format dates for emails (e.g., "Monday, January 15th")
function formatDateForEmail(dateStr) {
  try {
    var date = new Date(dateStr);
    if (isNaN(date.getTime())) {
      // Try parsing as YYYY-MM-DD
      var parts = dateStr.split('-');
      if (parts.length === 3) {
        date = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
      } else {
        return dateStr; // Return original if can't parse
      }
    }
    
    var days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    var months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    
    var dayName = days[date.getDay()];
    var monthName = months[date.getMonth()];
    var day = date.getDate();
    
    // Add ordinal suffix (1st, 2nd, 3rd, 4th, etc.)
    var suffix = 'th';
    if (day === 1 || day === 21 || day === 31) suffix = 'st';
    else if (day === 2 || day === 22) suffix = 'nd';
    else if (day === 3 || day === 23) suffix = 'rd';
    
    return dayName + ', ' + monthName + ' ' + day + suffix;
  } catch (e) {
    Logger.log('Error formatting date: ' + e.toString());
    return dateStr;
  }
}

// Helper function to format times for emails (e.g., "3:00 PM - 4:00 PM")
function formatTimeForEmail(timeStr) {
  try {
    // If it's already a formatted string like "3:00 PM - 4:00 PM", return it
    if (typeof timeStr === 'string' && timeStr.indexOf(' - ') !== -1 && (timeStr.indexOf('PM') !== -1 || timeStr.indexOf('AM') !== -1)) {
      return timeStr;
    }
    
    // Handle Date objects or ISO strings
    var date = null;
    if (timeStr instanceof Date) {
      date = timeStr;
    } else if (typeof timeStr === 'string') {
      // Try parsing as ISO string (e.g., "1899-12-30T15:00:00.000Z")
      if (timeStr.indexOf('T') !== -1 || timeStr.indexOf('Z') !== -1) {
        date = new Date(timeStr);
      } else {
        // If it's already a simple time string, return it
        return timeStr;
      }
    }
    
    if (date && !isNaN(date.getTime())) {
      // Extract hours and minutes from the date
      var hours = date.getHours();
      var minutes = date.getMinutes();
      var ampm = hours >= 12 ? 'PM' : 'AM';
      hours = hours % 12;
      hours = hours ? hours : 12; // the hour '0' should be '12'
      var minutesStr = minutes < 10 ? '0' + minutes : minutes;
      return hours + ':' + minutesStr + ' ' + ampm;
    }
    
    return timeStr;
  } catch (e) {
    Logger.log('Error formatting time: ' + e.toString());
    return timeStr;
  }
}

// Function to send approval emails
function sendApprovalEmails(submissionId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var subSh = ss.getSheetByName('Submissions');
    if (!subSh) return {success:false,error:'Submissions sheet not found'};
    
    // Find the submission
    var lr = subSh.getLastRow();
    if (lr < 2) return {success:false,error:'No submissions found'};
    
    var numCols = subSh.getLastColumn();
    var colCount = numCols >= 14 ? 14 : 13;
    var data = subSh.getRange(2,1,lr-1,colCount).getValues();
    
    var submission = null;
    var submissionIdStr = (submissionId + '').trim();
    
    for (var i = 0; i < data.length; i++) {
      var rowId = (data[i][0] + '').trim();
      if (rowId === submissionIdStr) {
        submission = {
          id: data[i][0],
          companyName: data[i][1] || '',
          email: data[i][2] || '',
          date: data[i][3] || '',
          startTime: data[i][5] || '',
          endTime: data[i][6] || ''
        };
        break;
      }
    }
    
    if (!submission) {
      Logger.log('❌ EMAIL: Submission not found: ' + submissionIdStr);
      return {success:false,error:'Submission not found'};
    }
    
    Logger.log('📧 EMAIL: Found submission for ' + submission.companyName);
    
    // Format date and time
    var formattedDate = formatDateForEmail(submission.date);
    var formattedStartTime = formatTimeForEmail(submission.startTime);
    var formattedEndTime = formatTimeForEmail(submission.endTime);
    var formattedTime = formattedStartTime + ' - ' + formattedEndTime;
    
    // Get or create "Approval Email" sheet
    var emailSh = ss.getSheetByName('Approval Email');
    if (!emailSh) {
      emailSh = ss.insertSheet('Approval Email');
      emailSh.getRange(1,1,1,4).setValues([['Email', '', '', 'Email Confirmation']]);
      Logger.log('✅ EMAIL: Created Approval Email sheet');
    }
    
    // Read Email 1 recipients from Column A (starting row 2)
    var email1Recipients = [];
    var email1LastRow = emailSh.getLastRow();
    if (email1LastRow >= 2) {
      var email1Data = emailSh.getRange(2, 1, email1LastRow - 1, 1).getValues();
      for (var j = 0; j < email1Data.length; j++) {
        var email = (email1Data[j][0] || '').toString().trim();
        if (email && email.indexOf('@') !== -1) {
          email1Recipients.push(email);
        }
      }
    }
    
    Logger.log('📧 EMAIL 1: Found ' + email1Recipients.length + ' recipients');
    
    // Read Email 2 CC recipients from Column D (starting row 2)
    var email2CC = [];
    var email2LastRow = emailSh.getLastRow();
    if (email2LastRow >= 2) {
      var email2CCData = emailSh.getRange(2, 4, email2LastRow - 1, 1).getValues();
      for (var k = 0; k < email2CCData.length; k++) {
        var ccEmail = (email2CCData[k][0] || '').toString().trim();
        if (ccEmail && ccEmail.indexOf('@') !== -1) {
          email2CC.push(ccEmail);
        }
      }
    }
    
    Logger.log('📧 EMAIL 2: Found ' + email2CC.length + ' CC recipients');
    
    // Email 1: To Ogilvy organizers
    if (email1Recipients.length > 0) {
      var email1Subject = 'Town Hall Booking Request for ' + submission.companyName + ' Presentation';
      var email1Body = 'Hey Lynn,\n\n' +
                      'Hope you\'re doing well! This is Tim\'s production minion reaching out to request a hold on the town hall from ' + 
                      formattedStartTime + ' to ' + formattedEndTime + ' on ' + formattedDate + '.\n\n' + 
                      submission.companyName + ' has been invited to come in and give a screening presentation, and we\'d love to lock in the space for that window.\n\n' +
                      'I\'m also emailing Tonya, Isaac, and Liv here for visibility — they\'ll take over the conversation once you\'re able to confirm availability.\n\n' +
                      'Hope you\'re having a lovely day!\n\n' +
                      'Best,\nTim\'s Production Minion';
      
      try {
        MailApp.sendEmail({
          to: email1Recipients.join(','),
          subject: email1Subject,
          body: email1Body
        });
        Logger.log('✅ EMAIL 1: Sent to ' + email1Recipients.length + ' recipients');
      } catch (e) {
        Logger.log('❌ EMAIL 1 ERROR: ' + e.toString());
      }
    } else {
      Logger.log('⚠️ EMAIL 1: No recipients found in Column A of Approval Email sheet');
    }
    
    // Email 2: To company with CC
    if (submission.email) {
      var email2Subject = 'Confirmation: Your Screening Time at WT3';
      var email2Body = 'Hi ' + submission.companyName + ' Team,\n\n' +
                      'This is a confirmation that your screening time has been locked in. Please review the details below:\n\n' +
                      '📍 Location: WT3 — 175 Greenwich St, 34th Floor New York, NY 10007\n\n' +
                      '🗓 Date & Time: ' + formattedDate + ', ' + formattedTime + ' (Please arrive 30 minutes early for setup and refreshments.)\n\n' +
                      '(We recommend to estimate for around 15-18 people this is subject to change and we would update you otherwise)\n\n' +
                      'Important Notes:\n\n' +
                      'Timing: All scheduled times are subject to change due to production needs or attendance adjustments. You\'ll be notified immediately if any updates occur.\n\n' +
                      'Food & Refreshments: If you plan to bring food, please confirm what you\'ll be providing as soon as possible so we can update the invites and get everyone excited. Our Ogilvy team will assist with receiving deliveries, but if they\'re in meetings when food arrives, please ensure your team is available to receive it.\n\n' +
                      'Attendance & Security: Please email your attendee list no later than 72 hours before your screening to: tonya.white@ogilvy.com\n' +
                      'CC: isaac.boruchowicz@ogilvy.com, liv.hatcher@ogilvy.com\n' +
                      'Include full names and the confirmed screening date for security clearance.\n\n' +
                      'We\'re excited to welcome you to WT3 and look forward to your presentation (and hopefully your food!).\n\n' +
                      'Best,\nTim\'s Production Team';
      
      try {
        var email2Options = {
          to: submission.email,
          subject: email2Subject,
          body: email2Body
        };
        
        if (email2CC.length > 0) {
          email2Options.cc = email2CC.join(',');
        }
        
        MailApp.sendEmail(email2Options);
        Logger.log('✅ EMAIL 2: Sent to ' + submission.email + ' with ' + email2CC.length + ' CC recipients');
      } catch (e) {
        Logger.log('❌ EMAIL 2 ERROR: ' + e.toString());
      }
    } else {
      Logger.log('⚠️ EMAIL 2: No company email found in submission');
    }
    
    return {success:true};
  } catch (e) {
    Logger.log('❌❌❌ EMAIL ERROR: ' + e.toString());
    return {success:false,error:e.toString()};
  }
}

// Function to send rejection email
function sendRejectionEmail(submissionId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var subSh = ss.getSheetByName('Submissions');
    if (!subSh) return {success:false,error:'Submissions sheet not found'};
    
    // Find the submission - check both Submissions and Rejections sheets
    var submission = null;
    var submissionIdStr = (submissionId + '').trim();
    
    // Try to find in Submissions first (in case it hasn't been moved yet)
    var lr = subSh.getLastRow();
    if (lr >= 2) {
      var numCols = subSh.getLastColumn();
      var colCount = numCols >= 14 ? 14 : 13;
      var data = subSh.getRange(2,1,lr-1,colCount).getValues();
      
      for (var i = 0; i < data.length; i++) {
        var rowId = (data[i][0] + '').trim();
        if (rowId === submissionIdStr) {
          submission = {
            id: data[i][0],
            companyName: data[i][1] || '',
            email: data[i][2] || ''
          };
          break;
        }
      }
    }
    
    // If not found in Submissions, try Rejections sheet
    if (!submission) {
      var rejSh = ss.getSheetByName('Rejections');
      if (rejSh) {
        var rejLr = rejSh.getLastRow();
        if (rejLr >= 2) {
          var rejColCount = rejSh.getLastColumn() >= 14 ? 14 : 13;
          var rejData = rejSh.getRange(2,1,rejLr-1,rejColCount).getValues();
          
          for (var j = 0; j < rejData.length; j++) {
            var rejRowId = (rejData[j][0] + '').trim();
            if (rejRowId === submissionIdStr) {
              submission = {
                id: rejData[j][0],
                companyName: rejData[j][1] || '',
                email: rejData[j][2] || ''
              };
              break;
            }
          }
        }
      }
    }
    
    if (!submission || !submission.email) {
      Logger.log('⚠️ REJECTION EMAIL: Submission not found or no email: ' + submissionIdStr);
      return {success:false,error:'Submission not found or no email'};
    }
    
    Logger.log('📧 REJECTION EMAIL: Found submission for ' + submission.companyName);
    
    // Get or create "Approval Email" sheet
    var emailSh = ss.getSheetByName('Approval Email');
    if (!emailSh) {
      emailSh = ss.insertSheet('Approval Email');
      emailSh.getRange(1,1,1,4).setValues([['Email', '', '', 'Email Confirmation']]);
      Logger.log('✅ REJECTION EMAIL: Created Approval Email sheet');
    }
    
    // Read CC recipients from Column D (starting row 2)
    var emailCC = [];
    var emailLastRow = emailSh.getLastRow();
    if (emailLastRow >= 2) {
      var emailCCData = emailSh.getRange(2, 4, emailLastRow - 1, 1).getValues();
      for (var k = 0; k < emailCCData.length; k++) {
        var ccEmail = (emailCCData[k][0] || '').toString().trim();
        if (ccEmail && ccEmail.indexOf('@') !== -1) {
          emailCC.push(ccEmail);
        }
      }
    }
    
    Logger.log('📧 REJECTION EMAIL: Found ' + emailCC.length + ' CC recipients');
    
    // Send rejection email
    var emailSubject = 'Screening Time Rejection Email for ' + submission.companyName;
    var emailBody = 'Hi ' + submission.companyName + ',\n\n' +
                    'Tim\'s production team has reviewed your submission and it has been rejected for the time you selected.\n\n' +
                    'I am adding the Ogilvy team here so you can coordinate directly on next steps and decide how to move forward.\n\n' +
                    'Thank you,\nIsaac';
    
    try {
      var emailOptions = {
        to: submission.email,
        subject: emailSubject,
        body: emailBody
      };
      
      if (emailCC.length > 0) {
        emailOptions.cc = emailCC.join(',');
      }
      
      MailApp.sendEmail(emailOptions);
      Logger.log('✅ REJECTION EMAIL: Sent to ' + submission.email + ' with ' + emailCC.length + ' CC recipients');
      return {success:true};
    } catch (e) {
      Logger.log('❌ REJECTION EMAIL ERROR: ' + e.toString());
      return {success:false,error:e.toString()};
    }
  } catch (e) {
    Logger.log('❌❌❌ REJECTION EMAIL ERROR: ' + e.toString());
    return {success:false,error:e.toString()};
  }
}

// Function to send time change email
function sendTimeChangeEmail(submissionId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var subSh = ss.getSheetByName('Submissions');
    if (!subSh) return {success:false,error:'Submissions sheet not found'};
    
    // Find the submission
    var lr = subSh.getLastRow();
    if (lr < 2) return {success:false,error:'No submissions found'};
    
    var numCols = subSh.getLastColumn();
    var colCount = numCols >= 14 ? 14 : 13;
    var data = subSh.getRange(2,1,lr-1,colCount).getValues();
    
    var submission = null;
    var submissionIdStr = (submissionId + '').trim();
    
    for (var i = 0; i < data.length; i++) {
      var rowId = (data[i][0] + '').trim();
      if (rowId === submissionIdStr) {
        submission = {
          id: data[i][0],
          companyName: data[i][1] || '',
          email: data[i][2] || '',
          date: data[i][3] || '',
          startTime: data[i][5] || '',
          endTime: data[i][6] || ''
        };
        break;
      }
    }
    
    if (!submission || !submission.email) {
      Logger.log('⚠️ TIME CHANGE EMAIL: Submission not found or no email: ' + submissionIdStr);
      return {success:false,error:'Submission not found or no email'};
    }
    
    Logger.log('📧 TIME CHANGE EMAIL: Found submission for ' + submission.companyName);
    
    // Format date and time
    var formattedDate = formatDateForEmail(submission.date);
    var formattedStartTime = formatTimeForEmail(submission.startTime);
    var formattedEndTime = formatTimeForEmail(submission.endTime);
    var formattedTime = formattedStartTime + ' - ' + formattedEndTime;
    
    // Get or create "Approval Email" sheet
    var emailSh = ss.getSheetByName('Approval Email');
    if (!emailSh) {
      emailSh = ss.insertSheet('Approval Email');
      emailSh.getRange(1,1,1,4).setValues([['Email', '', '', 'Email Confirmation']]);
      Logger.log('✅ TIME CHANGE EMAIL: Created Approval Email sheet');
    }
    
    // Read CC recipients from Column D (starting row 2)
    var emailCC = [];
    var emailLastRow = emailSh.getLastRow();
    if (emailLastRow >= 2) {
      var emailCCData = emailSh.getRange(2, 4, emailLastRow - 1, 1).getValues();
      for (var j = 0; j < emailCCData.length; j++) {
        var ccEmail = (emailCCData[j][0] || '').toString().trim();
        if (ccEmail && ccEmail.indexOf('@') !== -1) {
          emailCC.push(ccEmail);
        }
      }
    }
    
    Logger.log('📧 TIME CHANGE EMAIL: Found ' + emailCC.length + ' CC recipients');
    
    // Send time change email
    var emailSubject = 'Update: Your Screening Time at WT3 Has Been Changed';
    var emailBody = 'Hi ' + submission.companyName + ' Team,\n\n' +
                    'This is a notification that your screening time has been changed. Please review the updated details below:\n\n' +
                    '📍 Location: WT3 — 175 Greenwich St, 34th Floor New York, NY 10007\n\n' +
                    '🗓 Date & Time: ' + formattedDate + ', ' + formattedTime + ' (Please arrive 30 minutes early for setup and refreshments.)\n\n' +
                    'Important Notes:\n\n' +
                    'Timing: All scheduled times are subject to change due to production needs or attendance adjustments. You\'ll be notified immediately if any updates occur.\n\n' +
                    'Food & Refreshments: If you plan to bring food, please confirm what you\'ll be providing as soon as possible so we can update the invites and get everyone excited. Our Ogilvy team will assist with receiving deliveries, but if they\'re in meetings when food arrives, please ensure your team is available to receive it.\n\n' +
                    'Attendance & Security: Please email your attendee list no later than 72 hours before your screening to: tonya.white@ogilvy.com\n' +
                    'CC: isaac.boruchowicz@ogilvy.com, liv.hatcher@ogilvy.com\n' +
                    'Include full names and the confirmed screening date for security clearance.\n\n' +
                    'We\'re excited to welcome you to WT3 and look forward to your presentation (and hopefully your food!).\n\n' +
                    'Best,\nTim\'s Production Team';
    
    try {
      var emailOptions = {
        to: submission.email,
        subject: emailSubject,
        body: emailBody
      };
      
      if (emailCC.length > 0) {
        emailOptions.cc = emailCC.join(',');
      }
      
      MailApp.sendEmail(emailOptions);
      Logger.log('✅ TIME CHANGE EMAIL: Sent to ' + submission.email + ' with ' + emailCC.length + ' CC recipients');
      return {success:true};
    } catch (e) {
      Logger.log('❌ TIME CHANGE EMAIL ERROR: ' + e.toString());
      return {success:false,error:e.toString()};
    }
  } catch (e) {
    Logger.log('❌❌❌ TIME CHANGE EMAIL ERROR: ' + e.toString());
    return {success:false,error:e.toString()};
  }
}