const ExcelJS = require("exceljs");
const { fetchWithRetry } = require("./zohoClient");
const { formatDuration, formatGeneratedAt, secondsToSheetDuration } = require("./reportBuilder");

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const SHEET_TITLE = "Scheduled Work";
const HEADER_ROW = 6;
const COLUMN_COUNT = 11;

class MicrosoftGraphExcelClient {
  constructor(config) {
    this.tenantId = config.tenantId;
    this.clientId = config.clientId;
    this.clientSecret = config.clientSecret;
    this.driveId = config.driveId;
    this.folderPath = normalizeFolderPath(config.folderPath || "");
    this.companyName = config.companyName || "TPH Group";
    this.businessTimezone = config.businessTimezone;
    this.accessToken = null;
  }

  async createReportWorkbook({ title, report, reportDate }) {
    await this.runStage("Connecting to Microsoft Graph", () => this.refreshAccessToken());
    await this.runStage("Checking SharePoint destination", () => this.verifyDestinationFolder());

    console.log("Creating Excel workbook...");
    const buffer = await buildWorkbookBuffer({
      title,
      report,
      reportDate,
      companyName: this.companyName,
      businessTimezone: this.businessTimezone
    });
    console.log(`Excel workbook generated in memory (${formatBytes(buffer.length)}).`);

    console.log("Uploading workbook to SharePoint...");

    const fileName = `${sanitizeFileName(title)}.xlsx`;
    const uploadPath = encodeGraphPath(joinGraphPath(this.folderPath, fileName));
    const uploadUrl = `${GRAPH_BASE_URL}/drives/${encodeURIComponent(this.driveId)}/root:/${uploadPath}:/content?@microsoft.graph.conflictBehavior=rename`;

    const response = await fetchWithRetry(uploadUrl, {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${this.accessToken}`,
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      },
      body: buffer
    });

    if (!response.ok) {
      const message = await safeResponseText(response);
      throw new Error(`SharePoint upload failed for "${joinGraphPath(this.folderPath, fileName)}" (${response.status}): ${message}`);
    }

    const uploaded = await response.json();
    console.log(`SharePoint upload complete: ${uploaded.name || fileName}`);
    return {
      id: uploaded.id,
      name: uploaded.name,
      webUrl: uploaded.webUrl
    };
  }

  async refreshAccessToken() {
    const body = new URLSearchParams({
      client_id: this.clientId,
      client_secret: this.clientSecret,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials"
    });

    const response = await fetchWithRetry(`https://login.microsoftonline.com/${encodeURIComponent(this.tenantId)}/oauth2/v2.0/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body
    });

    if (!response.ok) {
      const message = await safeResponseText(response);
      throw new Error(`Microsoft Graph token request failed (${response.status}): ${message}`);
    }

    const payload = await response.json();
    if (!payload.access_token) {
      throw new Error("Microsoft Graph token request did not return an access token.");
    }

    this.accessToken = payload.access_token;
  }

  async verifyDestinationFolder() {
    const target = this.folderPath
      ? `/root:/${encodeGraphPath(this.folderPath)}`
      : "/root";
    const response = await fetchWithRetry(`${GRAPH_BASE_URL}/drives/${encodeURIComponent(this.driveId)}${target}`, {
      headers: {
        Authorization: `Bearer ${this.accessToken}`,
        Accept: "application/json"
      }
    });

    if (!response.ok) {
      const message = await safeResponseText(response);
      const folderLabel = this.folderPath || "/";
      throw new Error(`Cannot access SharePoint folder "${folderLabel}" in configured drive (${response.status}): ${message}`);
    }

    const folder = await response.json();
    if (!folder.folder && this.folderPath) {
      throw new Error(`Configured SharePoint path "${this.folderPath}" exists but is not a folder.`);
    }

    console.log(`SharePoint destination ready: ${folder.webUrl || this.folderPath || "drive root"}`);
  }

  async runStage(label, action) {
    console.log(`${label}...`);
    try {
      const result = await action();
      console.log(`${label} complete.`);
      return result;
    } catch (error) {
      throw new Error(`${label} failed: ${error.message}`);
    }
  }
}

async function buildWorkbookBuffer({ title, report, reportDate, companyName, businessTimezone }) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = companyName;
  workbook.created = new Date();
  workbook.modified = new Date();

  const worksheet = workbook.addWorksheet(SHEET_TITLE, {
    views: [{ state: "frozen", ySplit: HEADER_ROW }]
  });

  worksheet.columns = [
    { header: "", key: "groupLabel", width: 28 },
    { header: "", key: "resourceName", width: 28 },
    { header: "", key: "address", width: 32 },
    { header: "", key: "appointmentName", width: 18 },
    { header: "", key: "start", width: 18 },
    { header: "", key: "end", width: 18 },
    { header: "", key: "duration", width: 18 },
    { header: "", key: "status", width: 18 },
    { header: "", key: "workOrder", width: 22 },
    { header: "", key: "workOrderStatus", width: 20 },
    { header: "", key: "durationTotal", width: 24 }
  ];

  worksheet.mergeCells(`A1:${columnLetter(COLUMN_COUNT)}1`);
  worksheet.getCell("A1").value = "Technicians Scheduled Work";
  worksheet.getCell("A1").font = { bold: true, size: 14 };

  worksheet.mergeCells(`A2:${columnLetter(COLUMN_COUNT)}2`);
  worksheet.getCell("A2").value = `Generated by ${companyName} on ${formatGeneratedAt(businessTimezone)}`;

  worksheet.mergeCells(`A4:${columnLetter(COLUMN_COUNT)}4`);
  worksheet.getCell("A4").value = `Record Count : ${report.recordCount}    Sum of Scheduled Duration : ${formatDuration(report.totalSeconds)}`;

  const header = worksheet.getRow(HEADER_ROW);
  header.values = [
    "Service Resource Name",
    "Service Resource Name",
    "Service Address Name",
    "Appointment Name",
    "Scheduled Start Time",
    "Scheduled End Time",
    "Scheduled Duration",
    "Status",
    "Work Order",
    "Status (Work Order)",
    "Sum of Scheduled Duration"
  ];
  header.font = { bold: true };
  header.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE5E5E5" }
  };

  if (report.recordCount === 0) {
    worksheet.addRow([`No scheduled appointments found for ${reportDate}.`]);
  } else {
    for (const group of report.groups) {
      group.rows.forEach((row, index) => {
        const excelRow = worksheet.addRow([
          index === 0 ? `${group.serviceResourceName} ( ${group.rows.length} )` : "",
          row.serviceResourceName,
          row.serviceAddressName,
          row.appointmentName,
          row.scheduledStartDisplay,
          row.scheduledEndDisplay,
          secondsToSheetDuration(row.scheduledDurationSeconds),
          row.status,
          row.workOrderName,
          row.workOrderStatus,
          ""
        ]);
      });

      const subtotalRow = worksheet.addRow([
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        `Sum of Scheduled Duration for ${group.serviceResourceName}`,
        secondsToSheetDuration(group.totalSeconds)
      ]);
      subtotalRow.font = { bold: true };
    }
  }

  const grandTotalRow = worksheet.addRow([
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "Grand Total",
    secondsToSheetDuration(report.totalSeconds)
  ]);
  grandTotalRow.font = { bold: true };

  worksheet.eachRow((row) => {
    row.eachCell((cell) => {
      cell.alignment = { vertical: "top", wrapText: true };
    });
  });

  worksheet.autoFilter = {
    from: { row: HEADER_ROW, column: 1 },
    to: { row: HEADER_ROW, column: COLUMN_COUNT }
  };

  return workbook.xlsx.writeBuffer();
}

function columnLetter(columnNumber) {
  let dividend = columnNumber;
  let name = "";
  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    name = String.fromCharCode(65 + modulo) + name;
    dividend = Math.floor((dividend - modulo) / 26);
  }
  return name;
}

function normalizeFolderPath(value) {
  const trimmed = String(value || "").trim();
  if (!trimmed || trimmed === "/") return "";
  return trimmed.replace(/^\/+|\/+$/g, "");
}

function joinGraphPath(folderPath, fileName) {
  return folderPath ? `${folderPath}/${fileName}` : fileName;
}

function encodeGraphPath(path) {
  return path.split("/").map(encodeURIComponent).join("/");
}

function sanitizeFileName(value) {
  return String(value).replace(/[<>:"/\\|?*]/g, "-").trim();
}

function formatBytes(value) {
  if (!Number.isFinite(value)) return "unknown size";
  if (value < 1024) return `${value} bytes`;
  if (value < 1024 * 1024) return `${Math.round(value / 1024)} KB`;
  return `${(value / 1024 / 1024).toFixed(1)} MB`;
}

async function safeResponseText(response) {
  try {
    return await response.text();
  } catch (_error) {
    return "";
  }
}

module.exports = {
  MicrosoftGraphExcelClient
};
