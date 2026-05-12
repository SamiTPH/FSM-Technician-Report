require("dotenv").config({ quiet: true });

const { MicrosoftGraphExcelClient } = require("./microsoftGraphClient");
const {
  buildReport,
  filterAppointmentsForReportDate,
  getWorkOrderIds,
  nowInTimezoneDateKey
} = require("./reportBuilder");
const { ZohoClient } = require("./zohoClient");

async function main() {
  const config = loadConfig();
  const reportDate = resolveReportDate(config);
  const workbookTitle = `${config.reportTitlePrefix} - ${reportDate}`;

  console.log(`Building scheduled-work report for ${reportDate} (${config.businessTimezone}).`);
  console.log(`Report destination: Microsoft Graph drive ${maskId(config.microsoftDriveId)}, folder "${config.microsoftFolderPath || "/"}".`);

  const zoho = new ZohoClient({
    accountsUrl: config.zohoAccountsUrl,
    apiDomain: config.zohoFsmApiDomain,
    clientId: config.zohoClientId,
    clientSecret: config.zohoClientSecret,
    refreshToken: config.zohoRefreshToken
  });

  await runStage("Refreshing Zoho access token", () => zoho.refreshAccessToken());

  let usedFullScanFallback = false;
  const appointments = await runStage("Fetching Zoho Service Appointments for report date", async () => {
    try {
      const searched = await zoho.fetchServiceAppointmentsForDate(reportDate, config.businessTimezone);
      const searchMatches = filterAppointmentsForReportDate(
        searched,
        reportDate,
        config.businessTimezone
      );
      const fullScan = await zoho.fetchServiceAppointments();
      const fullScanMatches = filterAppointmentsForReportDate(
        fullScan,
        reportDate,
        config.businessTimezone
      );
      const merged = mergeAppointmentsById(searchMatches, fullScanMatches);
      const recoveredCount = merged.length - searchMatches.length;

      if (recoveredCount > 0) {
        usedFullScanFallback = true;
        console.warn(`Warning: Zoho date search missed ${recoveredCount} same-day appointment(s); recovered them with full scan fallback.`);
      }

      return merged;
    } catch (error) {
      console.warn(`Warning: date-filtered Zoho search failed; falling back to full appointment scan. ${error.message}`);
      usedFullScanFallback = true;
      return zoho.fetchServiceAppointments();
    }
  });
  const appointmentsForToday = filterAppointmentsForReportDate(
    appointments,
    reportDate,
    config.businessTimezone
  );
  console.log(`Fetched ${appointments.length} candidate appointments; ${appointmentsForToday.length} match ${reportDate}.`);
  if (usedFullScanFallback) {
    console.log("Full scan fallback was used to protect against Zoho search/reporting mismatches.");
  }

  const detailedAppointments = await fetchAppointmentDetails(zoho, appointmentsForToday);

  let users = [];
  try {
    console.log("Fetching Zoho users for service-resource lookup...");
    users = await zoho.fetchUsers();
    console.log(`Fetched ${users.length} Zoho users for service-resource lookup.`);
  } catch (error) {
    console.warn(`Warning: Users fetch failed; continuing with IDs or Unknown Resource. ${error.message}`);
  }

  const workOrdersById = new Map();
  const workOrderIds = getWorkOrderIds(detailedAppointments);
  console.log(`Fetching ${workOrderIds.length} linked Zoho Work Orders...`);
  for (const workOrderId of workOrderIds) {
    try {
      workOrdersById.set(workOrderId, await zoho.fetchWorkOrder(workOrderId));
    } catch (error) {
      console.warn(`Warning: Work Order ${workOrderId} fetch failed; continuing with blank details. ${error.message}`);
      workOrdersById.set(workOrderId, null);
    }
  }
  console.log("Linked Work Order fetch step complete.");

  console.log("Building grouped scheduled-work report...");
  const report = buildReport({
    appointments: detailedAppointments,
    users,
    workOrdersById,
    businessTimezone: config.businessTimezone,
    reportDate
  });
  console.log(`Report built: ${report.recordCount} rows, ${report.groups.length} service-resource groups.`);

  const microsoftExcel = new MicrosoftGraphExcelClient({
    tenantId: config.microsoftTenantId,
    clientId: config.microsoftClientId,
    clientSecret: config.microsoftClientSecret,
    driveId: config.microsoftDriveId,
    folderPath: config.microsoftFolderPath,
    companyName: config.companyName,
    businessTimezone: config.businessTimezone
  });

  const result = await microsoftExcel.createReportWorkbook({
    title: workbookTitle,
    report,
    reportDate
  });

  console.log(`Excel workbook created: ${result.webUrl}`);
}

async function runStage(label, action) {
  console.log(`${label}...`);
  try {
    const result = await action();
    console.log(`${label} complete.`);
    return result;
  } catch (error) {
    throw new Error(`${label} failed: ${error.message}`);
  }
}

async function fetchAppointmentDetails(zoho, appointments) {
  console.log(`Fetching detail records for ${appointments.length} Service Appointment(s)...`);
  const detailed = [];

  for (const appointment of appointments) {
    const appointmentId = appointment.id || appointment.ID;
    if (!appointmentId) {
      console.warn(`Warning: Appointment ${appointment.Name || appointment.name || "(unknown)"} has no ID; using shallow record.`);
      detailed.push(appointment);
      continue;
    }

    try {
      detailed.push(await zoho.fetchServiceAppointment(appointmentId));
    } catch (error) {
      console.warn(`Warning: Appointment detail fetch failed for ${appointment.Name || appointmentId}; using shallow record. ${error.message}`);
      detailed.push(appointment);
    }
  }

  console.log("Service Appointment detail fetch step complete.");
  return detailed;
}

function loadConfig() {
  const config = {
    zohoAccountsUrl: process.env.ZOHO_ACCOUNTS_URL,
    zohoFsmApiDomain: process.env.ZOHO_FSM_API_DOMAIN,
    zohoClientId: process.env.ZOHO_CLIENT_ID,
    zohoClientSecret: process.env.ZOHO_CLIENT_SECRET,
    zohoRefreshToken: process.env.ZOHO_REFRESH_TOKEN,
    businessTimezone: process.env.BUSINESS_TIMEZONE || "Asia/Dubai",
    microsoftTenantId: process.env.MICROSOFT_TENANT_ID,
    microsoftClientId: process.env.MICROSOFT_CLIENT_ID,
    microsoftClientSecret: process.env.MICROSOFT_CLIENT_SECRET,
    microsoftDriveId: process.env.MICROSOFT_DRIVE_ID,
    microsoftFolderPath: process.env.MICROSOFT_FOLDER_PATH || "",
    reportTitlePrefix: process.env.REPORT_TITLE_PREFIX || "Technicians Scheduled Work",
    companyName: process.env.COMPANY_NAME || "TPH Group",
    reportDateOverride: process.env.REPORT_DATE_OVERRIDE || "",
    allowReportDateOverride: process.env.ALLOW_REPORT_DATE_OVERRIDE === "true"
  };

  const required = [
    ["ZOHO_ACCOUNTS_URL", config.zohoAccountsUrl],
    ["ZOHO_FSM_API_DOMAIN", config.zohoFsmApiDomain],
    ["ZOHO_CLIENT_ID", config.zohoClientId],
    ["ZOHO_CLIENT_SECRET", config.zohoClientSecret],
    ["ZOHO_REFRESH_TOKEN", config.zohoRefreshToken],
    ["MICROSOFT_TENANT_ID", config.microsoftTenantId],
    ["MICROSOFT_CLIENT_ID", config.microsoftClientId],
    ["MICROSOFT_CLIENT_SECRET", config.microsoftClientSecret],
    ["MICROSOFT_DRIVE_ID", config.microsoftDriveId]
  ];

  const missing = required.filter(([, value]) => !value).map(([key]) => key);
  if (missing.length > 0) {
    throw new Error(`Missing required environment variables: ${missing.join(", ")}`);
  }

  return config;
}

function resolveReportDate(config) {
  if (config.reportDateOverride && config.allowReportDateOverride) {
    console.warn(`Using REPORT_DATE_OVERRIDE=${config.reportDateOverride}. Do not enable this in Railway production.`);
    return config.reportDateOverride;
  }

  if (config.reportDateOverride && !config.allowReportDateOverride) {
    console.warn("Ignoring REPORT_DATE_OVERRIDE because ALLOW_REPORT_DATE_OVERRIDE is not true.");
  }

  return nowInTimezoneDateKey(config.businessTimezone);
}

function maskId(value) {
  const text = String(value || "");
  if (text.length <= 10) return text || "(missing)";
  return `${text.slice(0, 5)}...${text.slice(-5)}`;
}

function mergeAppointmentsById(primary, secondary) {
  const merged = new Map();
  for (const appointment of [...primary, ...secondary]) {
    const key = appointment.id || appointment.ID || appointment.Name || JSON.stringify(appointment);
    if (!merged.has(key)) merged.set(key, appointment);
  }
  return [...merged.values()];
}

main().catch((error) => {
  console.error(`Run failed: ${error.message}`);
  process.exitCode = 1;
});
