import test from "node:test";
import assert from "node:assert/strict";
import { createLimitedCache } from "./fn-cache.js";

const tick = () => new Promise((r) => setTimeout(r, 0));

test("repeated keys are answered from cache, not re-run", async () => {
  const cache = createLimitedCache();
  let calls = 0;
  const task = async () => {
    calls++;
    return "hot";
  };
  assert.equal(await cache.run("k", task), "hot");
  assert.equal(await cache.run("k", task), "hot");
  assert.equal(await cache.run("k", task), "hot");
  assert.equal(calls, 1);
});

test("concurrent identical keys share one in-flight request", async () => {
  const cache = createLimitedCache();
  let calls = 0;
  const task = async () => {
    calls++;
    await tick();
    return "one";
  };
  const results = await Promise.all([cache.run("k", task), cache.run("k", task), cache.run("k", task)]);
  assert.deepEqual(results, ["one", "one", "one"]);
  assert.equal(calls, 1);
});

test("no more than maxConcurrent distinct tasks run at once", async () => {
  const cache = createLimitedCache({ maxConcurrent: 3 });
  let active = 0;
  let peak = 0;
  const task = async () => {
    active++;
    peak = Math.max(peak, active);
    await tick();
    active--;
    return "x";
  };
  // 20 distinct keys — this is the drag-fill case that used to fire 20 at once.
  await Promise.all(Array.from({ length: 20 }, (_, i) => cache.run(`k${i}`, task)));
  assert.equal(peak, 3);
});

test("every queued task still completes", async () => {
  const cache = createLimitedCache({ maxConcurrent: 2 });
  const out = await Promise.all(
    Array.from({ length: 10 }, (_, i) =>
      cache.run(`k${i}`, async () => {
        await tick();
        return i;
      })
    )
  );
  assert.deepEqual(out, [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
});

test("failures propagate and are not cached", async () => {
  const cache = createLimitedCache();
  let calls = 0;
  const flaky = async () => {
    calls++;
    if (calls === 1) throw new Error("provider down");
    return "recovered";
  };
  await assert.rejects(() => cache.run("k", flaky), /provider down/);
  // A cached error would pin the failure onto every cell until reload.
  assert.equal(await cache.run("k", flaky), "recovered");
  assert.equal(calls, 2);
});

test("a failing task releases its concurrency slot", async () => {
  const cache = createLimitedCache({ maxConcurrent: 1 });
  await assert.rejects(
    () =>
      cache.run("bad", async () => {
        throw new Error("nope");
      }),
    /nope/
  );
  // With maxConcurrent:1 this await would hang forever if the rejected task
  // had kept its slot — completing at all is the real proof.
  assert.equal(await cache.run("good", async () => "ok"), "ok");
  // The slot counter is decremented in a .finally(), i.e. one microtask after
  // the caller's promise settles, so let the queue drain before reading it.
  await tick();
  assert.equal(cache.stats.active, 0);
});

test("cache evicts oldest entries past maxEntries", async () => {
  const cache = createLimitedCache({ maxEntries: 3 });
  for (const k of ["a", "b", "c", "d"]) await cache.run(k, async () => k);
  assert.equal(cache.stats.cached, 3);
  // "a" was evicted, so it re-runs; "d" is still cached.
  let reran = false;
  await cache.run("a", async () => {
    reran = true;
    return "a";
  });
  assert.equal(reran, true);
  let dReran = false;
  await cache.run("d", async () => {
    dReran = true;
    return "d";
  });
  assert.equal(dReran, false);
});
