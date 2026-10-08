const STORAGE_KEY = "sandbox-signage:last-reliable-data";

function assertDisplayData(data) {
  if (!data || typeof data !== "object") throw new Error("Display data is not an object");
  if (!data.day || !Array.isArray(data.day.events)) throw new Error("Display data has no day schedule");
  if (!Array.isArray(data.workdays)) throw new Error("Display data has no workday overview");
  if (!data.ocean?.tide?.points?.length) throw new Error("Display data has no tide points");
  return data;
}

export async function fetchDisplayData({
  fetchImpl = fetch,
  storage = globalThis.localStorage,
  url = "/api/display-data",
} = {}) {
  try {
    const response = await fetchImpl(url, { cache: "no-store" });
    if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
    const data = assertDisplayData(await response.json());
    storage?.setItem(STORAGE_KEY, JSON.stringify(data));
    return data;
  } catch (error) {
    const cached = storage?.getItem(STORAGE_KEY);
    if (!cached) throw error;
    const data = assertDisplayData(JSON.parse(cached));
    return {
      ...data,
      health: {
        ...data.health,
        online: false,
        stale: true,
        message: `OFFLINE · LAST RELIABLE DATA ${data.health?.lastUpdated || "UNKNOWN"}`,
      },
    };
  }
}

export { STORAGE_KEY };
