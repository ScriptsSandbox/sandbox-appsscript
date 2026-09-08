import assert from "node:assert/strict";
import test from "node:test";
import { fetchDisplayData, STORAGE_KEY } from "../src/data-client.js";

function validData() {
  return {
    day: { events: [] },
    workdays: [],
    ocean: { tide: { points: [{ minute: 0, heightFt: 1 }] } },
    health: { stale: false, lastUpdated: "12:00" },
  };
}

function memoryStorage() {
  const values = new Map();
  return {
    getItem: (key) => values.get(key) ?? null,
    setItem: (key, value) => values.set(key, value),
  };
}

test("live display data is saved as the browser's last reliable copy", async () => {
  const storage = memoryStorage();
  const data = await fetchDisplayData({
    storage,
    fetchImpl: async () => ({ ok: true, json: async () => validData() }),
  });
  assert.equal(data.health.stale, false);
  assert.ok(storage.getItem(STORAGE_KEY));
});

test("browser falls back to its reliable copy when the local service is unreachable", async () => {
  const storage = memoryStorage();
  storage.setItem(STORAGE_KEY, JSON.stringify(validData()));
  const data = await fetchDisplayData({
    storage,
    fetchImpl: async () => { throw new Error("offline"); },
  });
  assert.equal(data.health.stale, true);
  assert.match(data.health.message, /OFFLINE/);
});
