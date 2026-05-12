const DATE_FIELD_CANDIDATES = [
  "Scheduled_Start_Date_Time",
  "Scheduled_Start",
  "Scheduled_Start_Time",
  "ScheduledStartDateTime",
  "scheduled_start_date_time"
];

const END_FIELD_CANDIDATES = [
  "Scheduled_End_Date_Time",
  "Scheduled_End",
  "Scheduled_End_Time",
  "ScheduledEndDateTime",
  "scheduled_end_date_time"
];

const APPOINTMENT_DATE_CANDIDATES = [
  "Appointment_Date",
  "Appointment_Date_Time",
  "Scheduled_Date",
  "Date"
];

const DURATION_CANDIDATES = [
  "Scheduled_Duration",
  "Scheduled_Duration_in_Minutes",
  "Scheduled_Duration_Minutes",
  "Duration",
  "Estimated_Duration"
];

const WORK_ORDER_CANDIDATES = [
  "Work_Order",
  "WorkOrder",
  "Work_Order_Name",
  "Parent_Work_Order"
];

const RESOURCE_FIELD_NAMES = [
  "$Service_Resources",
  "Service_Resources",
  "Service_Resource",
  "Service_Resource_Name",
  "Lead",
  "Crew",
  "Assigned_To",
  "Assigned_User",
  "Technician",
  "Technicians",
  "Resources"
];

function buildReport({ appointments, users, workOrdersById, businessTimezone, reportDate }) {
  const resourceLookup = buildServiceResourceLookup(users);
  const rows = [];
  const rowKeys = new Set();

  for (const appointment of appointments) {
    const scheduledStart = getFirst(appointment, DATE_FIELD_CANDIDATES);
    const scheduledEnd = getFirst(appointment, END_FIELD_CANDIDATES);
    const displayMode = getDateDisplayMode(scheduledStart, scheduledEnd, reportDate, businessTimezone);
    const durationSeconds = getScheduledDurationSeconds(appointment, scheduledStart, scheduledEnd);
    const workOrderLookup = getLookup(appointment, WORK_ORDER_CANDIDATES);
    const workOrder = workOrderLookup.id ? workOrdersById.get(workOrderLookup.id) : null;
    const serviceResources = extractServiceResources(appointment, resourceLookup);
    const expandedResources = serviceResources.length
      ? serviceResources
      : [{ id: "", name: "Unassigned" }];

    for (const resource of expandedResources) {
      const serviceResourceName = resource.name || resourceLookup.get(resource.id) || resource.id || "Unknown Resource";
      const rowKey = [
        getId(appointment),
        normalizeKey(serviceResourceName),
        scheduledStart,
        scheduledEnd,
        workOrderLookup.id || ""
      ].join("|");
      if (rowKeys.has(rowKey)) continue;
      rowKeys.add(rowKey);

      rows.push({
        appointmentId: getId(appointment),
        appointmentName: getName(appointment),
        serviceResourceId: resource.id || "",
        serviceResourceName,
        serviceAddressName: extractServiceAddress(appointment, workOrder),
        scheduledStart,
        scheduledEnd,
        scheduledStartDisplay: formatZohoDateTime(scheduledStart, businessTimezone, displayMode),
        scheduledEndDisplay: formatZohoDateTime(scheduledEnd, businessTimezone, displayMode),
        scheduledDurationSeconds: durationSeconds,
        status: getStatus(appointment),
        workOrderName: workOrderLookup.name || extractWorkOrderName(workOrder),
        workOrderStatus: extractWorkOrderStatus(workOrder)
      });
    }
  }

  rows.sort((a, b) => {
    const byResource = a.serviceResourceName.localeCompare(b.serviceResourceName);
    if (byResource !== 0) return byResource;
    return dateValue(a.scheduledStart) - dateValue(b.scheduledStart);
  });

  const groups = [];
  for (const row of rows) {
    let group = groups[groups.length - 1];
    if (!group || group.serviceResourceName !== row.serviceResourceName) {
      group = {
        serviceResourceName: row.serviceResourceName,
        rows: [],
        totalSeconds: 0
      };
      groups.push(group);
    }
    group.rows.push(row);
    group.totalSeconds += row.scheduledDurationSeconds || 0;
  }

  const totalSeconds = rows.reduce((sum, row) => sum + (row.scheduledDurationSeconds || 0), 0);

  return {
    reportDate,
    recordCount: rows.length,
    totalSeconds,
    groups
  };
}

function filterAppointmentsForReportDate(appointments, reportDate, businessTimezone) {
  return appointments.filter((appointment) => {
    const scheduledStart = getFirst(appointment, DATE_FIELD_CANDIDATES);
    const scheduledEnd = getFirst(appointment, END_FIELD_CANDIDATES);
    if (scheduledEnd && scheduledEndMatchesReportDate(scheduledEnd, reportDate, businessTimezone)) {
      return true;
    }

    if (!scheduledStart && !scheduledEnd) {
      for (const field of APPOINTMENT_DATE_CANDIDATES) {
        const value = appointment[field];
        if (value && singleDateMatchesReportDate(value, reportDate, businessTimezone)) {
          return true;
        }
      }
    }

    return false;
  });
}

function buildServiceResourceLookup(users) {
  const lookup = new Map();

  for (const user of users || []) {
    collectResourceLookupEntries(user, lookup);
  }

  return lookup;
}

function collectResourceLookupEntries(value, lookup) {
  if (!value || typeof value !== "object") return;

  if (Array.isArray(value)) {
    for (const item of value) collectResourceLookupEntries(item, lookup);
    return;
  }

  if (isResourceLike(value)) {
    const id = String(value.id || value.ID || value.Service_Resource_ID || value.resource_id || "");
    const name = getName(value);
    if (id && name) lookup.set(id, name);
  }

  for (const [key, child] of Object.entries(value)) {
    if (/service[_\s-]?resources?/i.test(key)) {
      collectResourceLookupEntries(child, lookup);
    }
  }
}

function extractServiceResources(appointment, resourceLookup) {
  const resources = [];

  for (const fieldName of RESOURCE_FIELD_NAMES) {
    collectResourcesFromValue(appointment[fieldName], resources, resourceLookup);
  }

  for (const [key, value] of Object.entries(appointment)) {
    if (/resource|technician|crew|assigned|lead/i.test(key)) {
      collectResourcesFromValue(value, resources, resourceLookup);
    }
  }

  const deduped = new Map();
  for (const resource of resources) {
    const key = resource.name ? `name:${normalizeKey(resource.name)}` : `id:${resource.id}`;
    if (key && !deduped.has(key)) deduped.set(key, resource);
  }

  return [...deduped.values()];
}

function collectResourcesFromValue(value, resources, resourceLookup) {
  if (!value) return;

  if (Array.isArray(value)) {
    for (const item of value) collectResourcesFromValue(item, resources, resourceLookup);
    return;
  }

  if (typeof value === "string" || typeof value === "number") {
    const text = String(value);
    resources.push({ id: resourceLookup.has(text) ? text : "", name: resourceLookup.get(text) || text });
    return;
  }

  if (typeof value !== "object") return;

  const id = String(value.id || value.ID || value.Service_Resource_ID || value.resource_id || "");
  const name = getName(value) || (id ? resourceLookup.get(id) : "");
  if (id || name) {
    resources.push({ id, name: name || id });
  }
}

function isResourceLike(value) {
  return Boolean(value.id || value.ID || value.Service_Resource_ID || value.resource_id) && Boolean(getName(value));
}

function getWorkOrderIds(appointments) {
  const ids = new Set();
  for (const appointment of appointments) {
    const lookup = getLookup(appointment, WORK_ORDER_CANDIDATES);
    if (lookup.id) ids.add(lookup.id);
  }
  return [...ids];
}

function getLookup(record, fieldNames) {
  for (const fieldName of fieldNames) {
    const value = record[fieldName];
    const parsed = parseLookup(value);
    if (parsed.id || parsed.name) return parsed;
  }
  return {};
}

function parseLookup(value) {
  if (!value) return {};
  if (typeof value === "string" || typeof value === "number") return { id: String(value), name: "" };
  if (typeof value !== "object") return {};
  return {
    id: String(value.id || value.ID || value.value || ""),
    name: getName(value)
  };
}

function extractServiceAddress(appointment, workOrder) {
  return stringifyAddress(getFirst(workOrder || {}, [
    "Service_Address",
    "Service_Address_Name",
    "Service_Address1",
    "Service_Site",
    "Site",
    "Site_Name",
    "Address",
    "Company",
    "Customer",
    "Contact"
  ], true)) || stringifyAddress(getFirst(appointment, [
    "Service_Address",
    "Service_Address_Name",
    "Service_Site",
    "Site",
    "Site_Name",
    "Address",
    "Company",
    "Customer"
  ], true)) || "";
}

function extractWorkOrderName(workOrder) {
  return getName(workOrder || {});
}

function extractWorkOrderStatus(workOrder) {
  return getStatus(workOrder || {});
}

function getStatus(record) {
  return stringifyLookup(getFirst(record, ["Status", "status", "Appointment_Status", "Work_Order_Status"], true));
}

function getName(record) {
  return stringifyLookup(getFirst(record || {}, ["name", "Name", "Full_Name", "Service_Resource_Name", "Work_Order_Name", "Appointment_Name"], true));
}

function getId(record) {
  return String(getFirst(record || {}, ["id", "ID", "Service_Appointment_ID"], false) || "");
}

function getFirst(record, fieldNames, allowObject = false) {
  for (const fieldName of fieldNames) {
    if (record && record[fieldName] !== undefined && record[fieldName] !== null && record[fieldName] !== "") {
      const value = record[fieldName];
      return allowObject ? value : stringifyLookup(value);
    }
  }
  return "";
}

function stringifyLookup(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return String(value);
  if (typeof value === "object") {
    return value.name || value.Name || value.display_value || value.Display_Value || value.value || value.id || "";
  }
  return "";
}

function stringifyAddress(value) {
  if (!value) return "";
  if (typeof value === "string" || typeof value === "number") return String(value);
  if (typeof value !== "object") return "";

  return value.Service_Address_Name
    || value.name
    || value.Name
    || value.Site_Name
    || value.Address_Name
    || value.Service_Street_1
    || "";
}

function normalizeKey(value) {
  return String(value || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function getScheduledDurationSeconds(appointment, scheduledStart, scheduledEnd) {
  for (const fieldName of DURATION_CANDIDATES) {
    const parsed = parseDurationToSeconds(appointment[fieldName]);
    if (parsed > 0) return parsed;
  }

  if (scheduledStart && scheduledEnd) {
    const start = new Date(scheduledStart);
    const end = new Date(scheduledEnd);
    if (!Number.isNaN(start.getTime()) && !Number.isNaN(end.getTime()) && end > start) {
      return Math.round((end.getTime() - start.getTime()) / 1000);
    }
  }

  return 0;
}

function parseDurationToSeconds(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") {
    return value > 24 ? Math.round(value * 60) : Math.round(value * 3600);
  }

  const text = stringifyLookup(value).trim();
  if (!text) return 0;

  const hhmmss = text.match(/^(\d+):([0-5]?\d)(?::([0-5]?\d))?$/);
  if (hhmmss) {
    return Number(hhmmss[1]) * 3600 + Number(hhmmss[2]) * 60 + Number(hhmmss[3] || 0);
  }

  const decimal = Number(text);
  if (Number.isFinite(decimal)) return parseDurationToSeconds(decimal);

  let seconds = 0;
  const hours = text.match(/(\d+(?:\.\d+)?)\s*(h|hr|hour)/i);
  const minutes = text.match(/(\d+(?:\.\d+)?)\s*(m|min|minute)/i);
  if (hours) seconds += Number(hours[1]) * 3600;
  if (minutes) seconds += Number(minutes[1]) * 60;
  return Math.round(seconds);
}

function formatDuration(seconds) {
  const total = Math.max(0, Math.round(seconds || 0));
  const hours = total / 3600;
  if (Number.isInteger(hours)) {
    return `${hours} ${hours === 1 ? "hr" : "hrs"}`;
  }

  const rounded = Math.round(hours * 100) / 100;
  return `${rounded} hrs`;
}

function secondsToSheetDuration(seconds) {
  return formatDuration(seconds);
}

function formatZohoDateTime(value, timeZone, displayMode = "business") {
  if (!value) return "";

  if (displayMode === "raw") {
    return formatRawZohoDateTime(value);
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);

  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone,
    day: "2-digit",
    month: "2-digit",
    year: "2-digit",
    hour: "numeric",
    minute: "2-digit",
    hour12: false
  }).formatToParts(date);
  const map = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${map.day}-${map.month}-${map.year} ${map.hour}:${map.minute}`;
}

function formatRawZohoDateTime(value) {
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
  if (!match) return String(value);
  const [, year, month, day, hour, minute] = match;
  return `${day}-${month}-${year.slice(2)} ${Number(hour)}:${minute}`;
}

function scheduledRangeMatchesReportDate(startValue, endValue, reportDate, timeZone) {
  if (localRangeTouchesReportDate(startValue, endValue, reportDate, timeZone)) return true;

  return rawRangeMatchesReportDate(startValue, endValue, reportDate);
}

function scheduledEndMatchesReportDate(endValue, reportDate, timeZone) {
  return localDateKey(endValue, timeZone) === reportDate || rawDateKey(endValue) === reportDate;
}

function singleDateMatchesReportDate(value, reportDate, timeZone) {
  return localDateKey(value, timeZone) === reportDate || rawDateKey(value) === reportDate;
}

function getDateDisplayMode(startValue, endValue, reportDate, timeZone) {
  const localTouches = startValue && localRangeTouchesReportDate(startValue, endValue, reportDate, timeZone);
  if (startValue && rawStartedBeforeAndEndsToday(startValue, endValue, reportDate)) return "raw";
  if (startValue && rawRangeMatchesReportDate(startValue, endValue, reportDate) && !localTouches) return "raw";
  if (localTouches) return "business";
  return "business";
}

function localRangeTouchesReportDate(startValue, endValue, reportDate, timeZone) {
  const localStart = localDateKey(startValue, timeZone);
  const localEnd = endValue ? localDateKey(endValue, timeZone) : localStart;
  if (!localStart) return false;

  return localStart <= reportDate && localEnd >= reportDate;
}

function rawRangeMatchesReportDate(startValue, endValue, reportDate) {
  const rawStart = rawDateKey(startValue);
  const rawEnd = endValue ? rawDateKey(endValue) : rawStart;
  if (!rawStart) return false;

  const startsAndEndsToday = rawStart === reportDate && rawEnd === reportDate;
  return startsAndEndsToday || rawStartedBeforeAndEndsToday(startValue, endValue, reportDate);
}

function rawStartedBeforeAndEndsToday(startValue, endValue, reportDate) {
  const rawStart = rawDateKey(startValue);
  const rawEnd = endValue ? rawDateKey(endValue) : rawStart;
  return Boolean(rawStart) && rawStart < reportDate && rawEnd === reportDate;
}

function rawDateKey(value) {
  const match = String(value || "").match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : "";
}

function localDateKey(value, timeZone) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    const match = String(value).match(/^(\d{4}-\d{2}-\d{2})/);
    return match ? match[1] : "";
  }

  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(date);
  const map = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${map.year}-${map.month}-${map.day}`;
}

function nowInTimezoneDateKey(timeZone) {
  return localDateKey(new Date().toISOString(), timeZone);
}

function formatGeneratedAt(timeZone) {
  const datePart = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).format(new Date());
  const timePart = new Intl.DateTimeFormat("en-US", {
    timeZone,
    hour: "numeric",
    minute: "2-digit",
    hour12: true
  }).format(new Date());
  return `${datePart} ${timePart}`;
}

function dateValue(value) {
  const timestamp = new Date(value).getTime();
  return Number.isNaN(timestamp) ? 0 : timestamp;
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

module.exports = {
  buildReport,
  filterAppointmentsForReportDate,
  formatDuration,
  formatGeneratedAt,
  getWorkOrderIds,
  nowInTimezoneDateKey,
  secondsToSheetDuration
};
