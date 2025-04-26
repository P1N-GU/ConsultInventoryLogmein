function checkApiRequest() {
  // Company ID collected in Logmein panel
  const companyId = "INSERT YOUR COMPANY ID";
  //API key generated in Logmein panel
  const psk = "INSERT YOUR API KEY HERE";
  const base64Auth = Utilities.base64Encode(companyId + ":" + psk);
  const groupUrl = "https://secure.logmein.com/public-api/v2/hostswithgroups";
  const inventoryReportUrl = "https://secure.logmein.com/public-api/v1/inventory/system/reports";
  const inventoryReportHardwareUrl = "https://secure.logmein.com/public-api/v1/inventory/hardware/reports";
  
  // Insert bellow the group for consult in Logmein
  const nomeGrupo = "INSERT GROUP NAME HERE";

  //To send 'nomeGrupo' as a parameter to the configured the  spreadsheet
  const { sheet, sheetUrl } = setupSpreadsheet(nomeGrupo); // Configure the spreadsheet and get the object and URL

  const options = {
    method: "GET",
    headers: { "Authorization": "Basic " + base64Auth },
    muteHttpExceptions: true
  };

  try {
    const groupResponse = UrlFetchApp.fetch(groupUrl, options);
    if (groupResponse.getResponseCode() === 200) {
      Logger.log("The API request was successful");
      const responseData = JSON.parse(groupResponse.getContentText());
      const groupMap = new Map();
      if (responseData.groups && Array.isArray(responseData.groups)) {
        responseData.groups.forEach(group => {
          groupMap.set(group.id, group.name || "Unnamed");
        });
      }

      // Filter the hosts what belongs to the specific group
      const filteredHosts = responseData.hosts.filter(host => groupMap.get(host.groupid) === nomeGrupo);

      if (filteredHosts.length > 0) {
        filteredHosts.forEach(host => {
          const groupName = groupMap.get(host.groupid) || "Unknown Group";
          const reportBody = {
            hostIds: [host.id],
            fields: ["LastLogonUserName", "OsType", "LastBootDate"]
          };
          const reportHardwareBody = {
            hostIds: [host.id],
            fields: ["HardwareModel", "ServiceTag", "CpuType", "MemorySize", "DriveCapacity"]
          };

          // Creates and retrieves the system and hardware inventory report
          const inventoryOptions = {
            method: "POST",
            headers: {
              "Authorization": "Basic " + base64Auth,
              "Content-Type": "application/json"
            },
            payload: JSON.stringify(reportBody),
            muteHttpExceptions: true
          };
          const reportToken = createInventoryReport(inventoryReportUrl, inventoryOptions, host.id);

          let reportData = {};
          if (reportToken) {
            const reportUrl = `https://secure.logmein.com/public-api/v1/inventory/system/reports/${reportToken}`;
            const reportOptions = {
              method: "GET",
              headers: { "Authorization": "Basic " + base64Auth },
              muteHttpExceptions: true
            };
            const reportResponse = UrlFetchApp.fetch(reportUrl, reportOptions);
            if (reportResponse.getResponseCode() === 200) {
              reportData = JSON.parse(reportResponse.getContentText());
              Logger.log("System inventory data:");
            } else {
              Logger.log("Error retrieving the system inventory report.");
            }
          }

          // Creates and retrievs the hardware inventory report
          const inventoryOptionsHardware = {
            method: "POST",
            headers: {
              "Authorization": "Basic " + base64Auth,
              "Content-Type": "application/json"
            },
            payload: JSON.stringify(reportHardwareBody),
            muteHttpExceptions: true
          };
          const reportHardwareToken = createInventoryReport(inventoryReportHardwareUrl, inventoryOptionsHardware, host.id);

          let reportHardwareData = {};
          if (reportHardwareToken) {
            const reportHardwareUrl = `https://secure.logmein.com/public-api/v1/inventory/hardware/reports/${reportHardwareToken}`;
            const reportHardwareOptions = {
              method: "GET",
              headers: { "Authorization": "Basic " + base64Auth },
              muteHttpExceptions: true
            };
            const reportHardwareResponse = UrlFetchApp.fetch(reportHardwareUrl, reportHardwareOptions);
            if (reportHardwareResponse.getResponseCode() === 200) {
              reportHardwareData = JSON.parse(reportHardwareResponse.getContentText());
              Logger.log("Hardware inventory data:");
            } else {
              Logger.log("Error ritrieving the hardware inventory report.");
            }
          }

          // Export data for spreadsheet, if the data are available
          if (reportData.hosts && reportHardwareData.hosts) {
            exportToSheet(sheet, groupName, host.description, host.id, reportData, reportHardwareData);
          } else {
            Logger.log("Data inventory absent or incomplete.");
          }
        });
      } else {
        Logger.log("No hosts found in the specific group.");
      }
    } else {
      Logger.log("Request error: " + groupResponse.getContentText());
    }
  } catch (error) {
    Logger.log("Error making request: " + error.toString());
  }
  Logger.log('Spreadsheet created: ' + sheetUrl);
}


//Function to configure the spreadsheet
function setupSpreadsheet(nomeGrupo) {
  const today = new Date();
  const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM-dd-yyyy');
  const spreadsheetName = `LogMeIn Inventory - ${nomeGrupo} - ${formattedDate}`; 
  const spreadsheet = SpreadsheetApp.create(spreadsheetName);
  const sheet = spreadsheet.getActiveSheet();
  sheet.setName('Devices');
  sheet.appendRow(["Group", "Computer Name", "ID Host", "Model", "ServiceTAG", "Last Logged in User", "CPU", "RAM", "Storage", "Operating System", "Last Initilization"]);
  return { sheet, sheetUrl: spreadsheet.getUrl() };
}

// Function to export data to the spreadsheet
function exportToSheet(sheet, groupName, hostDescription, hostId, reportData, reportHardwareData) {
  const hardwareData = reportHardwareData.hosts ? reportHardwareData.hosts[hostId] : {};

  // Correct the access to 'hardwareModel' and 'serviceTag'
  const hardwareModel = hardwareData.hardwareInfo?.model || "Unknown";
  const serviceTag = hardwareData.serviceTag || "Unknown";

  // Adjust to correctly get the 'driveCapacity' from the primary drive (if it exists)
  const driveCapacity = hardwareData.drives && hardwareData.drives.length > 0
    ? hardwareData.drives[0].capacity || "Unknown"
    : "Unknown";

  const lastLogonUserName = reportData.hosts[hostId]?.lastLogonUserName || "N/A";
  const cpuType = hardwareData.processors && hardwareData.processors.length > 0
    ? hardwareData.processors[0].type || "Unknown"
    : "Unknown";
  const memorySize = hardwareData.memories && hardwareData.memories.length > 0
    ? hardwareData.memories[0].size || "Unknown"
    : "Unknown";
  const osType = reportData.hosts[hostId]?.operatingSystem?.type || "Unknown";
  const lastBootDate = reportData.hosts[hostId]?.lastBootDate || "N/A";

  // Add a new line with the data to the spreadsheet
  sheet.appendRow([groupName, hostDescription, hostId, hardwareModel, serviceTag, lastLogonUserName, cpuType, memorySize, driveCapacity, osType, lastBootDate]);
}


// Helper function to create the inventory report with limited attempts
function createInventoryReport(url, options, hostId, retries = 3) {
  for (let i = 0; i < retries; i++) {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    if (responseCode === 200 || responseCode === 201) {
      const data = JSON.parse(response.getContentText());
      return data.token;  // Return the report token if successful
    } else if (responseCode === 429) {
      Logger.log("Error 429: Request limit exceeded. Waiting before trying again...");
      Utilities.sleep(60000); // Wait 60 seconds before trying again
    } else {
      Logger.log("Error requesting the inventory report for the host" + hostId);
      Logger.log("Response code: " + responseCode);
      Logger.log("Response content: " + response.getContentText());
      return null;
    }
  }
  Logger.log("Failed to create the inventory report for the host after multiple attempts " + hostId);
  return null;
}
