const DEFAULT_PER_PAGE = 200;

class ZohoClient {
  constructor(config) {
    this.accountsUrl = trimTrailingSlash(config.accountsUrl);
    this.apiDomain = trimTrailingSlash(config.apiDomain);
    this.clientId = config.clientId;
    this.clientSecret = config.clientSecret;
    this.refreshToken = config.refreshToken;
    this.accessToken = null;
  }

  async refreshAccessToken() {
    const body = new URLSearchParams({
      grant_type: "refresh_token",
      client_id: this.clientId,
      client_secret: this.clientSecret,
      refresh_token: this.refreshToken
    });

    const response = await fetchWithRetry(`${this.accountsUrl}/oauth/v2/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body
    });

    if (!response.ok) {
      const message = await safeResponseText(response);
      throw new Error(`Zoho token refresh failed (${response.status}): ${message}`);
    }

    const payload = await response.json();
    if (!payload.access_token) {
      throw new Error("Zoho token refresh did not return an access token.");
    }

    this.accessToken = payload.access_token;
    return this.accessToken;
  }

  async fetchServiceAppointments() {
    return this.fetchPaginated("/fsm/v1/Service_Appointments");
  }

  async fetchServiceAppointmentsForDate(reportDate, businessTimezone) {
    const range = buildZohoDateRange(reportDate, businessTimezone);
    const params = new URLSearchParams({
      api_name: "Scheduled_End_Date_Time",
      comparator: "between",
      value: `${range.start},${range.end}`
    });

    console.log(`Zoho Service Appointments scheduled-end search range: ${range.start} to ${range.end}.`);
    return this.fetchPaginated(`/fsm/v1/Service_Appointments/search?${params.toString()}`);
  }

  async fetchServiceAppointment(serviceAppointmentId) {
    const payload = await this.requestJson(`/fsm/v1/Service_Appointments/${encodeURIComponent(serviceAppointmentId)}`);
    const records = normalizeRecords(payload);
    return records[0] || payload.data || payload;
  }

  async fetchUsers() {
    const payload = await this.requestJson("/fsm/v1/users");
    return normalizeRecords(payload);
  }

  async fetchAttendance() {
    return this.fetchPaginated("/fsm/v1/Attendance");
  }

  async fetchServiceResources() {
    return this.fetchPaginated("/fsm/v1/Service_Resources");
  }

  async fetchWorkOrder(workOrderId) {
    const payload = await this.requestJson(`/fsm/v1/Work_Orders/${encodeURIComponent(workOrderId)}`);
    const records = normalizeRecords(payload);
    return records[0] || payload.data || payload;
  }

  async fetchPaginated(path) {
    const all = [];
    let page = 1;
    let hasMore = true;

    while (hasMore) {
      const separator = path.includes("?") ? "&" : "?";
      const payload = await this.requestJson(`${path}${separator}page=${page}&per_page=${DEFAULT_PER_PAGE}`);
      all.push(...normalizeRecords(payload));

      const info = payload.info || {};
      hasMore = Boolean(info.more_records || info.moreRecords);

      if (!hasMore && Array.isArray(payload.data) && payload.data.length === DEFAULT_PER_PAGE) {
        hasMore = true;
      }

      page += 1;
    }

    return all;
  }

  async requestJson(path) {
    if (!this.accessToken) {
      throw new Error("Zoho access token is missing. Call refreshAccessToken() first.");
    }

    const response = await fetchWithRetry(`${this.apiDomain}${path}`, {
      headers: {
        Authorization: `Zoho-oauthtoken ${this.accessToken}`,
        Accept: "application/json"
      }
    });

    if (!response.ok) {
      const message = await safeResponseText(response);
      throw new Error(`Zoho request failed (${response.status}) for ${path}: ${message}`);
    }

    return response.json();
  }
}

async function fetchWithRetry(url, options = {}, maxAttempts = 4) {
  let lastError;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      const response = await fetch(url, options);
      if (!shouldRetry(response.status) || attempt === maxAttempts) {
        return response;
      }

      const retryAfter = Number(response.headers.get("retry-after"));
      const delay = Number.isFinite(retryAfter) && retryAfter > 0
        ? retryAfter * 1000
        : 500 * 2 ** (attempt - 1);
      await sleep(delay);
    } catch (error) {
      lastError = error;
      if (attempt === maxAttempts) {
        throw new Error(`Network request failed for ${redactUrl(url)} after ${maxAttempts} attempts: ${error.message}`);
      }
      await sleep(500 * 2 ** (attempt - 1));
    }
  }

  throw lastError || new Error("Request failed.");
}

function shouldRetry(status) {
  return status === 429 || status >= 500;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function safeResponseText(response) {
  try {
    return await response.text();
  } catch (_error) {
    return "";
  }
}

function normalizeRecords(payload) {
  if (Array.isArray(payload)) return payload;
  if (Array.isArray(payload.data)) return payload.data;
  if (Array.isArray(payload.users)) return payload.users;
  if (Array.isArray(payload.Service_Appointments)) return payload.Service_Appointments;
  if (Array.isArray(payload.Work_Orders)) return payload.Work_Orders;
  if (Array.isArray(payload.Attendance)) return payload.Attendance;
  if (Array.isArray(payload.Service_Resources)) return payload.Service_Resources;
  return [];
}

function trimTrailingSlash(value) {
  return String(value || "").replace(/\/+$/, "");
}

function redactUrl(url) {
  try {
    const parsed = new URL(url);
    parsed.search = "";
    return parsed.toString();
  } catch (_error) {
    return "request URL";
  }
}

function buildZohoDateRange(reportDate, timeZone) {
  const offset = getOffsetStringForDate(reportDate, timeZone);
  return {
    start: `${reportDate}T00:00:00${offset}`,
    end: `${reportDate}T23:59:59${offset}`
  };
}

function getOffsetStringForDate(reportDate, timeZone) {
  const [year, month, day] = reportDate.split("-").map(Number);
  const probe = new Date(Date.UTC(year, month - 1, day, 12, 0, 0));
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone,
    timeZoneName: "longOffset"
  }).formatToParts(probe);
  const zoneName = parts.find((part) => part.type === "timeZoneName")?.value || "GMT+00:00";
  const match = zoneName.match(/GMT([+-]\d{2}):?(\d{2})?/);
  if (!match) return "+00:00";
  return `${match[1]}:${match[2] || "00"}`;
}

module.exports = {
  ZohoClient,
  fetchWithRetry
};
