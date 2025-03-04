/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Bind encryption and decryption button event handlers
    const encryptButton = document.getElementById("encrypt");
    const decryptButton = document.getElementById("decrypt");
    
    if (encryptButton) {
      encryptButton.onclick = encryptSelectedColumn;
    }
    
    if (decryptButton) {
      decryptButton.onclick = decryptSelectedColumn;
    }
    
    // Set up password dialog event handlers
    const dialog = document.getElementById("passwordDialog");
    const confirmButton = document.getElementById("confirmPassword");
    const cancelButton = document.getElementById("cancelPassword");
    
    if (confirmButton) {
      confirmButton.onclick = () => {
        const password = document.getElementById("password").value;
        if (password) {
          if (dialog.dataset.action === 'decrypt') {
            processDecryption(password);
          } else {
            processEncryption(password);
          }
          dialog.close();
        }
      };
    }
    
    if (cancelButton) {
      cancelButton.onclick = () => {
        dialog.close();
      };
    }
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

async function encryptSelectedColumn() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["columnIndex", "values"]);
      await context.sync();
      
      // Check if only one column is selected
      if (range.columnIndex === undefined) {
        throw new Error("Please select a column");
      }
      
      // Show password input dialog
      document.getElementById("passwordDialog").showModal();
    });
  } catch (error) {
    console.error(error);
    // Display error message
    Office.context.ui.displayDialogAsync(`Error: ${error.message}`);
  }
}

async function decryptSelectedColumn() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["columnIndex", "values"]);
      await context.sync();
      
      // Check if only one column is selected
      if (range.columnIndex === undefined) {
        throw new Error("Please select a column");
      }
      
      // Get current worksheet
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Get the row count of used range
      const usedRange = sheet.getUsedRange();
      usedRange.load(["rowCount"]);
      await context.sync();
      
      // Get entire column data
      const columnIndex = range.columnIndex;
      const wholeColumn = sheet.getRangeByIndexes(0, columnIndex, usedRange.rowCount, 1);
      wholeColumn.load(["values", "rowCount"]);
      
      // Load header cell format
      const headerCell = wholeColumn.getCell(0, 0);
      headerCell.format.load("fill");
      await context.sync();

      // Save header content
      const headerText = wholeColumn.values[0][0];
      console.log('Header content before decryption:', headerText);

      // Check if encrypted (by checking header cell background color)
      console.log('Header cell background color:', headerCell.format.fill.color);
      const headerColor = headerCell.format.fill.color;
      if (!headerColor || headerColor.toLowerCase() !== "#c8e6c9") {
        console.log('This column is not encrypted');
        await Office.context.ui.displayDialogAsync(
          "This column is not encrypted",
          {height: 30, width: 30}
        );
        return;
      }

      // Show password input dialog
      const dialog = document.getElementById("passwordDialog");
      dialog.dataset.action = 'decrypt';  // Mark current operation as decrypt
      dialog.showModal();
    });
  } catch (error) {
    console.error(error);
    Office.context.ui.displayDialogAsync(`Error: ${error.message}`);
  }
}

async function processEncryption(password) {
  try {
    await Excel.run(async (context) => {
      // Get selected range
      const range = context.workbook.getSelectedRange();
      range.load(["columnCount", "rowCount", "values", "columnIndex"]);
      await context.sync();

      // Check if only one column is selected
      if (range.columnCount !== 1) {
        throw new Error("Please select only one column");
      }

      console.log('Selected range data:', range.values);
      console.log('Selected column index:', range.columnIndex);

      // Get current worksheet
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Get row count of used range
      const usedRange = sheet.getUsedRange();
      usedRange.load(["rowCount", "address"]);
      await context.sync();
      
      console.log('Used range row count:', usedRange.rowCount);
      
      // Get entire column data (from first row to last row of used range)
      const columnIndex = range.columnIndex;  // Using 0-based index
      const wholeColumn = sheet.getRangeByIndexes(0, columnIndex, usedRange.rowCount, 1);
      wholeColumn.load(["values", "rowCount"]);
      
      // Load header cell
      const headerCell = wholeColumn.getCell(0, 0);
      headerCell.format.load("fill");
      await context.sync();

      // Check if already encrypted (by checking header cell background color)
      console.log('Header cell background color:', headerCell.format.fill.color);
      if (headerCell.format.fill.color === "#C8E6C9") {
        // Display dialog with message
        await Office.context.ui.displayDialogAsync(
          "This column is already encrypted",
          {height: 30, width: 30}
        );
        return;
      }

      if (!wholeColumn.values || wholeColumn.values.length === 0) {
        throw new Error("Unable to read column data");
      }

      console.log('Entire column data:', wholeColumn.values);

      // Extract all non-empty data except header
      const dataToEncrypt = wholeColumn.values
        .slice(1)  // Exclude header row
        .map(row => row[0])  // Get first (and only) cell of each row
        .filter(cell => cell !== "" && cell !== null && cell !== undefined);  // Filter empty values

      if (dataToEncrypt.length === 0) {
        throw new Error("No valid data to encrypt in selected column");
      }

      console.log('Data to encrypt:', dataToEncrypt);

      // Record original row numbers (for later updates)
      const validDataRows = wholeColumn.values
        .slice(1)  // Exclude header row
        .map((row, index) => ({ 
          value: row[0], 
          rowIndex: index + 1  // +1 because we skipped header row
        }))
        .filter(item => item.value !== "" && item.value !== null && item.value !== undefined);

      console.log('Valid data rows:', validDataRows);

      // Base64 encode password
      const base64Password = btoa(password);
      console.log('Base64 encoded password:', base64Password);
      
      // Build auth header
      const authHeader = `VSAuth vsauth_method="sharedSecret",vsauth_data="${base64Password}",vsauth_identity_ascii="demo@voltage.com",vsauth_version="200"`;
      console.log('Complete auth header:', authHeader);
      
      // Prepare request data
      const requestBody = {
        format: "AUTO",
        data: dataToEncrypt
      };

      console.log('=== API Request Start ===');
      const apiUrl = 'https://voltage-pp-0000.dataprotection.voltage.com/vibesimple/rest/v1/protect';
      console.log('Request URL:', apiUrl);
      console.log('Request method:', 'POST');
      console.log('Request headers:', {
        'Content-Type': 'application/json',
        'Authorization': authHeader
      });
      console.log('Request body:', JSON.stringify(requestBody, null, 2));
      
      const startTime = new Date();
      
      try {
        // Add timeout setting
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 30000); // 30 seconds timeout

        // Call encryption API
        const response = await fetch(apiUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': authHeader
          },
          body: JSON.stringify(requestBody),
          signal: controller.signal,
          redirect: 'follow',
          referrerPolicy: 'no-referrer'
        });

        clearTimeout(timeoutId);

        const endTime = new Date();
        const duration = endTime - startTime;

        console.log('=== API Response Info ===');
        console.log('Response status:', response.status, response.statusText);
        console.log('Response headers:', Object.fromEntries([...response.headers]));
        console.log('Request duration:', duration, 'ms');

        if (!response.ok) {
          const errorText = await response.text();
          console.error('Error response body:', errorText);
          throw new Error(`Encryption service request failed: ${response.status} ${response.statusText} - ${errorText}`);
        }

        const encryptedData = await response.json();
        console.log('Response body:', JSON.stringify(encryptedData, null, 2));
        
        // Validate response data format
        if (!encryptedData || !encryptedData.data || !Array.isArray(encryptedData.data)) {
          throw new Error('Invalid server response format');
        }

        // Verify returned data length matches sent data length
        if (encryptedData.data.length !== dataToEncrypt.length) {
          throw new Error('Encrypted data length does not match original data');
        }

        console.log('=== API Request End ===');
        
        // Update Excel data (only update non-empty cells)
        validDataRows.forEach((item, index) => {
          const cell = wholeColumn.getCell(item.rowIndex, 0);
          cell.values = [[encryptedData.data[index]]];
          cell.format.fill.color = "#E8F5E9";  // Add background color for encrypted data
        });
        
        // Set header cell style
        const titleCell = wholeColumn.getCell(0, 0);
        titleCell.format.fill.color = "#C8E6C9";  // Slightly darker green for header
        
        // Ensure worksheet protection is enabled
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.protection.load("protected");
        await context.sync();
        
        // First remove worksheet protection (if exists)
        if (worksheet.protection.protected) {
            worksheet.protection.unprotect();
            await context.sync();
        }
        
        // Get entire used range
        const entireRange = worksheet.getUsedRange();
        entireRange.load("columnCount");
        await context.sync();
        
        // First unlock all cells
        entireRange.format.protection.locked = false;
        await context.sync();
        
        // Lock all cells in encrypted column
        for (let i = 0; i < wholeColumn.rowCount; i++) {
            const cell = wholeColumn.getCell(i, 0);
            cell.format.protection.locked = true;
        }
        await context.sync();
        
        // Re-enable worksheet protection, but allow other operations
        worksheet.protection.protect({
            allowInsertRows: true,
            allowInsertColumns: true,
            allowDeleteRows: true,
            allowDeleteColumns: true,
            allowSort: true,
            allowFilter: true,
            allowEditObjects: true,
            allowEditScenarios: true
        });
        await context.sync();
        
        // Add comment
        const commentText = `🔒 Encryption time: ${new Date().toLocaleString()}\nEncrypted data count: ${encryptedData.data.length}`;
        titleCell.worksheet.comments.add(titleCell, commentText);
      } catch (fetchError) {
        console.error('=== Network Request Error ===');
        console.error('Error type:', fetchError.name);
        console.error('Error message:', fetchError.message);
        
        if (fetchError.name === 'AbortError') {
          throw new Error('Request timeout, please check your network connection');
        } else if (fetchError.message.includes('Failed to fetch')) {
          throw new Error('Unable to connect to encryption server, please check:\n1. Company network connection\n2. VPN status');
        } else {
          throw fetchError;
        }
      }
    });
  } catch (error) {
    console.error('=== API Error ===');
    console.error('Error details:', error);
    console.error('Error stack:', error.stack);
    
    // Provide friendly error message
    let errorMessage = 'Error during encryption';
    if (error.message.includes('Cannot read properties of null')) {
      errorMessage = 'Invalid server response format';
    } else if (error.message.includes('Failed to fetch')) {
      errorMessage = error.message;
    } else if (error.message.includes('timeout')) {
      errorMessage = 'Request timeout, please check network connection';
    } else {
      errorMessage = `Encryption failed: ${error.message}`;
    }
    
    Office.context.ui.displayDialogAsync(errorMessage);
  }
}

async function processDecryption(password) {
  try {
    await Excel.run(async (context) => {
      // Get selected range
      const range = context.workbook.getSelectedRange();
      range.load(["columnCount", "rowCount", "values", "columnIndex"]);
      await context.sync();

      // Check if only one column is selected
      if (range.columnCount !== 1) {
        throw new Error("Please select only one column");
      }

      // Get current worksheet
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Get row count of used range
      const usedRange = sheet.getUsedRange();
      usedRange.load(["rowCount"]);
      await context.sync();
      
      // Get entire column data
      const columnIndex = range.columnIndex;
      const wholeColumn = sheet.getRangeByIndexes(0, columnIndex, usedRange.rowCount, 1);
      wholeColumn.load(["values", "rowCount"]);
      
      // Load header cell format
      const headerCell = wholeColumn.getCell(0, 0);
      headerCell.format.load("fill");
      await context.sync();

      // Save header content
      const headerText = wholeColumn.values[0][0];
      console.log('Header content before decryption:', headerText);

      // Check if encrypted (by checking header cell background color)
      console.log('Header cell background color:', headerCell.format.fill.color);
      const headerColor = headerCell.format.fill.color;
      if (!headerColor || headerColor.toLowerCase() !== "#c8e6c9") {
        console.log('This column is not encrypted');
        await Office.context.ui.displayDialogAsync(
          "This column is not encrypted",
          {height: 30, width: 30}
        );
        return;
      }

      // Extract all non-empty data except header
      const dataToDecrypt = wholeColumn.values
        .slice(1)  // Exclude header row
        .map(row => row[0])  // Get first cell of each row
        .filter(cell => cell !== "" && cell !== null && cell !== undefined);  // Filter empty values

      if (dataToDecrypt.length === 0) {
        throw new Error("No valid data to decrypt in selected column");
      }

      // Record original row numbers
      const validDataRows = wholeColumn.values
        .slice(1)
        .map((row, index) => ({ 
          value: row[0], 
          rowIndex: index + 1
        }))
        .filter(item => item.value !== "" && item.value !== null && item.value !== undefined);

      // Base64 encode password
      const base64Password = btoa(password);
      
      // Build auth header
      const authHeader = `VSAuth vsauth_method="sharedSecret",vsauth_data="${base64Password}",vsauth_identity_ascii="demo@voltage.com",vsauth_version="200"`;
      
      // Prepare request data
      const requestBody = {
        format: "AUTO",
        data: dataToDecrypt
      };

      console.log('=== Decryption API Request Start ===');
      const apiUrl = 'https://voltage-pp-0000.dataprotection.voltage.com/vibesimple/rest/v1/access';
      
      try {
        // Add timeout setting
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 30000);

        // Call decryption API
        const response = await fetch(apiUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': authHeader
          },
          body: JSON.stringify(requestBody),
          signal: controller.signal
        });

        clearTimeout(timeoutId);

        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`Decryption service request failed: ${response.status} ${response.statusText} - ${errorText}`);
        }

        const decryptedData = await response.json();
        
        if (!decryptedData || !decryptedData.data || !Array.isArray(decryptedData.data)) {
          throw new Error('Invalid server response format');
        }

        // First remove worksheet protection
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.protection.load("protected");
        await context.sync();
        
        if (worksheet.protection.protected) {
          worksheet.protection.unprotect();
          await context.sync();
        }

        // Update Excel data
        validDataRows.forEach((item, index) => {
          const cell = wholeColumn.getCell(item.rowIndex, 0);
          cell.values = [[decryptedData.data[index]]];
          cell.format.fill.clear();  // Clear background color
        });

        // Clear header cell style and comments, but keep content
        const titleCell = wholeColumn.getCell(0, 0);
        titleCell.format.fill.clear();  // Clear background color
        titleCell.clear(Excel.ClearApplyTo.comments);  // Clear comments
        
        // Rewrite header content
        await context.sync();
        titleCell.values = [[headerText]];
        await context.sync();
        
        console.log('Header content after decryption:', headerText);

        // Show success message
        await Office.context.ui.displayDialogAsync(
          "Decryption completed!",
          {height: 30, width: 30}
        );

      } catch (fetchError) {
        console.error('=== Network Request Error ===');
        if (fetchError.name === 'AbortError') {
          throw new Error('Request timeout, please check your network connection');
        } else if (fetchError.message.includes('Failed to fetch')) {
          throw new Error('Unable to connect to decryption server, please check:\n1. Company network connection\n2. VPN status');
        } else {
          throw fetchError;
        }
      }
    });
  } catch (error) {
    console.error('=== Decryption Error ===');
    console.error('Error details:', error);
    
    let errorMessage = 'Error during decryption';
    if (error.message.includes('Cannot read properties of null')) {
      errorMessage = 'Invalid server response format';
    } else if (error.message.includes('Failed to fetch')) {
      errorMessage = error.message;
    } else if (error.message.includes('timeout')) {
      errorMessage = 'Request timeout, please check network connection';
    } else {
      errorMessage = `Decryption failed: ${error.message}`;
    }
    
    Office.context.ui.displayDialogAsync(errorMessage);
  }
}
