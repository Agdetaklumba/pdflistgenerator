function generateDeliveryPDFs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var deliveriesSheet = ss.getSheetByName("Daily Deliveries");
  var deliveryListSheet = ss.getSheetByName("Delivery List");
  var returnListSheet = ss.getSheetByName("Return List");

  var lastRowA = deliveriesSheet.getRange("A:A").getValues().filter(String).length + 7;
  var lastRowE = deliveriesSheet.getRange("E:E").getValues().filter(String).length + 7;
  var maxRow = Math.max(lastRowA, lastRowE);

  var deliveryIds = deliveriesSheet.getRange("A8:A" + lastRowA).getValues();
  var returnIds = deliveriesSheet.getRange("E8:E" + lastRowE).getValues();
  var deliveryCarriers = deliveriesSheet.getRange("C8:C" + lastRowA).getValues();
  var returnCarriers = deliveriesSheet.getRange("G8:G" + lastRowE).getValues();
  var deliveryorderIds = deliveriesSheet.getRange("B8:B" + lastRowA).getValues();
  var returnorderIds = deliveriesSheet.getRange("F8:F" + lastRowE).getValues();

  var dateString = deliveriesSheet.getRange("B1").getDisplayValue();
  var mainFolderName = "Delivery Lists";
  var mainFolder = getOrCreateFolder(mainFolderName);
  var dateFolderName = dateString || "New Folder";
  var dateFolder = getOrCreateFolder(dateFolderName, mainFolder);

  var allPdfUrls = [];
  var carrierGroups = {};
  var orderIdsString = "";
  deliveriesSheet.getRange("D8:D20").clear();
  deliveriesSheet.getRange("H8:H20").clear();

  for (var i = 0; i < maxRow - 7; i++) {
    var deliveryId = (i < deliveryIds.length && deliveryIds[i] && deliveryIds[i][0]) ? deliveryIds[i][0] : "";
    var returnId = (i < returnIds.length && returnIds[i] && returnIds[i][0]) ? returnIds[i][0] : "";
    var deliveryorderId = (deliveryorderIds[i] && deliveryorderIds[i][0]) ? deliveryorderIds[i][0] : "";
    var returnorderId = (returnorderIds[i] && returnorderIds[i][0]) ? returnorderIds[i][0] : "";
    var deliveryCarrier = (deliveryCarriers[i] && deliveryCarriers[i][0]) ? deliveryCarriers[i][0] : "";
    var returnCarrier = (returnCarriers[i] && returnCarriers[i][0]) ? returnCarriers[i][0] : "";

    // Handle Delivery
    if (deliveryId && deliveryCarrier) {
      if (!carrierGroups[deliveryCarrier]) {
        carrierGroups[deliveryCarrier] = { deliveries: [], returns: [] };
      }
      deliveryListSheet.getRange("B6").setValue(deliveryId);
      SpreadsheetApp.flush();
      Utilities.sleep(2500); // Wait for calculations to complete
      var deliveryPdfBlob = exportSheetToPDF(deliveryListSheet);
      var deliveryFile = dateFolder.createFile(deliveryPdfBlob.setName("DeliveryList_" + deliveryorderId + ".pdf"));
      var deliveryDownloadUrl = generateDirectDownloadUrl(deliveryFile.getId());
      allPdfUrls.push(deliveryDownloadUrl);
      carrierGroups[deliveryCarrier].deliveries.push(deliveryDownloadUrl);
      setLinkInSheet(deliveryDownloadUrl, "D" + (i + 8));
      if (orderIdsString.length > 0) {
        orderIdsString += ", "; // Add comma only if it's not the first ID
      }
      orderIdsString += deliveryorderId;
    }

    // Handle Return
    if (returnId && returnCarrier) {
      if (!carrierGroups[returnCarrier]) {
        carrierGroups[returnCarrier] = { deliveries: [], returns: [] };
      }
      returnListSheet.getRange("B6").setValue(returnId);
      SpreadsheetApp.flush();
      Utilities.sleep(2500); // Wait for calculations to complete
      var returnPdfBlob = exportSheetToPDF(returnListSheet);
      var returnFile = dateFolder.createFile(returnPdfBlob.setName("ReturnList_" + returnorderId + ".pdf"));
      var returnDownloadUrl = generateDirectDownloadUrl(returnFile.getId());
      allPdfUrls.push(returnDownloadUrl);
      carrierGroups[returnCarrier].returns.push(returnDownloadUrl);
      setLinkInSheet(returnDownloadUrl,"H"+(i+8));
      if (orderIdsString.length > 0) {
        orderIdsString += ", ";
      }
      orderIdsString += returnorderId;
    }
  }
  deliveriesSheet.getRange("D1:D6").clear();
  // Process each carrier group
  for (var carrier in carrierGroups) {
    var carrierData = carrierGroups[carrier];
    var combinedPdfUrls = carrierData.deliveries.concat(carrierData.returns);
    Logger.log(combinedPdfUrls);

    if (combinedPdfUrls.length > 0) {
      var carrierPdfUrl;
      // Check if there's only one PDF in the group
      if (combinedPdfUrls.length === 1) {
        // Use the single PDF URL directly
        carrierPdfUrl = combinedPdfUrls[0];
      } else {
        // Merge multiple PDFs
        carrierPdfUrl = mergePDFsUsingCloudConvert(combinedPdfUrls, dateFolder, carrier);
      }
      displayLinkInSheet(carrierPdfUrl, "D" + (Object.keys(carrierGroups).indexOf(carrier) + 2));
    }
  }
  Logger.log(carrierGroups);
  Logger.log(orderIdsString);
}



function displayLinkInSheet(url, cellRef) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Daily Deliveries");
  var cell = sheet.getRange(cellRef);

  cell.setFormula('=HYPERLINK("' + url + '", "Click here to open PDF")');
}



function getOrCreateFolder(folderName, parentFolder) {
  var folder;
  var folders = (parentFolder ? parentFolder : DriveApp).getFoldersByName(folderName);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = (parentFolder ? parentFolder : DriveApp).createFolder(folderName);
  }
  return folder;
}

function exportSheetToPDF(sheet) {
  var spreadsheet = sheet.getParent();
  var sheetId = sheet.getSheetId();
  var url = spreadsheet.getUrl().replace(/\/edit.*$/, '');
  var pdfOptions = '/export?exportFormat=pdf&format=pdf' + // Export format
                   '&size=letter' + // Paper size
                   '&portrait=true' + // Orientation, false for landscape
                   '&fitw=true' + // Fit to width, false for actual size
                   '&sheetnames=false&printtitle=false' + // Headers and footers options
                   '&pagenumbers=true' + // Page numbers
                   '&gridlines=false' + // Gridlines
                   '&fzr=false' + // Repeat frozen rows
                   '&gid=' + sheetId; // Sheet ID

  var headers = {
    'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()
  };

  var response = UrlFetchApp.fetch(url + pdfOptions, {headers: headers});
  return response.getBlob().setName(sheet.getName() + '.pdf');
}

function mergePDFsUsingCloudConvert(pdfUrls, dateFolder, orderIdsString) {
  var apiKey = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiMDkwN2ExZjdhMGM2ZGQyNDRkNjI1MTczYjU1ZTJmOTUyYjJkOTM3OWUyNzI5YjY0NWM1ZGI3MzlhYjRiZGMwYWRmZmI0ZTkxMjVjMmQ0ZmYiLCJpYXQiOjE3MDI3NTA2OTIuOTcwMzk5LCJuYmYiOjE3MDI3NTA2OTIuOTcwNCwiZXhwIjo0ODU4NDI0MjkyLjk2NjMzLCJzdWIiOiI2NjUzMjg2NiIsInNjb3BlcyI6WyJ1c2VyLnJlYWQiLCJ1c2VyLndyaXRlIiwidGFzay53cml0ZSIsInRhc2sucmVhZCIsIndlYmhvb2sucmVhZCIsInByZXNldC5yZWFkIiwid2ViaG9vay53cml0ZSIsInByZXNldC53cml0ZSJdfQ.YjTGTqpFdFs1FzAXfJkYrQd58w1HooMu9wUEavXV0cCPm5Lu5LYwYcVcgjKpEVGIsx79oW05LvGvH6uDgsRgi1JE9tJDEbV9STlzrNB-Ony811uwxAHfCcEXif44kAq1658P3vXrFIQI6qLJj0kRjoFllqeFR-9Wq9A1bdfaOUOf-oILisdFGV5LnuAqybMZoAJnIwilohXbDscHUOkAZQXviZ0vfjpQemQn7LpMGoB4LCC2A1lJIwrqO2JkEn3peEYu76u-7rQRYbc2Z4aW4XDDk70QZ18EEK0pFHJCaU6xGXBSs1k3Kcl8YhgizMJ1zCw5aP4MUSHryjvY7ycENp4odKWUX7MHQt6z_XZlB8XsrvhybQtpEwuJyl3WHDnW4jCtKnkcmMlf6JdvCuqB9naBDO8ikaoiFijJGznNwhzQF2EgCU9MSPGLBLCdcjj123_cQkO0kwKu6AOVQPvnBkuBsYWCDWwdJs3DPRXGNnMQOb2QLb6t-I5AxWWhAFm9K3iMnlbGkleXt2f1B6XpHeGg5RgKD9fE4CLS5q2Sks0dJJwM_nG5txaA0MYuyoOYCrbHh1Ql2w-n-8qa7cKIx5xBFOYUQZ7_VgXKfiNq0CRjyHgrbsESqa03Zm8aLPUOr0m5owGz-kH8tnkxa-n_9ecOX-HuKDnZqYjM-FZbH3Q'; // Replace with your CloudConvert API key

  // Define tasks for importing files and merging them
  var tasks = {};
  pdfUrls.forEach(function(url, index) {
    var importTaskName = 'import-file-' + index;
    tasks[importTaskName] = {
      'operation': 'import/url',
      'url': url
    };
  });

  // Merge task
  var mergeTaskName = 'merge';
  tasks[mergeTaskName] = {
    'operation': 'merge',
    'input': Object.keys(tasks),
    'output_format': 'pdf'  // Specify the output format here
  };


  // Export task
  var exportTaskName = 'export-result';
  tasks[exportTaskName] = {
    'operation': 'export/url',
    'input': mergeTaskName
  };

  var jobPayload = { tasks: tasks };
  
  // Create the job in CloudConvert
  var jobResponse = createCloudConvertJob(apiKey, jobPayload);

  // Handle the job's result
  if (jobResponse && jobResponse.data && jobResponse.data.id) {
    // Wait for the job to finish and get the download URL
    var downloadUrl = waitForCloudConvertJob(apiKey, jobResponse.data.id, exportTaskName);
    if (downloadUrl) {
      // Download the merged PDF and save it to Google Drive
      var urltoreturn = downloadAndSaveFile(downloadUrl, dateFolder, orderIdsString);
    }
  }
  return urltoreturn
}

function createCloudConvertJob(apiKey, jobPayload) {
  var url = 'https://api.cloudconvert.com/v2/jobs';
  var options = {
    'method' : 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    },
    'payload' : JSON.stringify(jobPayload)
  };

  var response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}

function waitForCloudConvertJob(apiKey, jobId, exportTaskName) {
  var checkUrl = 'https://api.cloudconvert.com/v2/jobs/' + jobId;
  var options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    }
  };
  var statusResponse;
  var downloadUrl = null;

  // Polling loop
  do {
    Utilities.sleep(5000); // Wait for 5 seconds before each check
    statusResponse = UrlFetchApp.fetch(checkUrl, options);
    var jobStatus = JSON.parse(statusResponse.getContentText());

    if (jobStatus.data.status === 'finished') {
      downloadUrl = jobStatus.data.tasks.find(function(task) {
        return task.name === exportTaskName && task.result && task.result.files;
      }).result.files[0].url;
      break;
    }
  } while (statusResponse && jobStatus.data.status !== 'error');

  return downloadUrl;
}

function downloadAndSaveFile(downloadUrl, dateFolder,orderIdsString) {
  var response = UrlFetchApp.fetch(downloadUrl);
  var responseBlob = response.getBlob();

  // Get today's date in the desired format
  // Create a new Date object for the current date and time
  var today = new Date();

  // Add one day to the date
  var tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);

  // Format the date to "yyyyMMdd"
  var formattedDate = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), "yyyyMMdd");
  var mergedPdfName = formattedDate + ": " + orderIdsString + ".pdf";
 // Naming the file with today's date

  // Save the merged PDF to the dateFolder with the new name
  var mergedPdfFile = dateFolder.createFile(responseBlob.setName(mergedPdfName));
  Logger.log(mergedPdfFile.getUrl());
  return mergedPdfFile.getUrl(); // Log the URL of the merged PDF
}

function generateDirectDownloadUrl(fileId) {
  // Construct the direct download URL for a file stored in Google Drive
  return "https://drive.google.com/uc?export=download&id=" + fileId;
}

function setLinkInSheet(url, cellRef) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Daily Deliveries");
  var cell = sheet.getRange(cellRef);

  cell.setFormula('=HYPERLINK("' + url + '", "Open PDF")');
}
