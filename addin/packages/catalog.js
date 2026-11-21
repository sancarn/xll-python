// Parsing the full PyPI project list and matching names happen off the UI thread.
const indexUrl = "https://pypi.org/simple/";
let names, loading, latest;

async function download(cache) {
  const response = await fetch(indexUrl, { headers: { Accept: "application/vnd.pypi.simple.v1+json" } });
  if (!response.ok) throw Error("PyPI is unavailable.");
  const projects = (await response.json()).projects.map(p => p.name);
  if (cache) {
    // Store only the searchable names. Cache storage can be unavailable or full.
    await cache.put("/names", new Response(projects.join("\n"), {
      headers: { "X-Saved": String(Date.now()) }
    })).catch(() => {});
  }
  return projects;
}

async function load() {
  const cache = await caches.open("pypi-project-names-v1").catch(() => null);
  const cached = cache && await cache.match("/names");
  if (!cached) return download(cache);
  const saved = (await cached.text()).split("\n");
  if (Date.now() - Number(cached.headers.get("X-Saved")) > 86400000) {
    download(cache).then(fresh => { names = fresh; search(); }).catch(() => {});
  }
  return saved;
}

const key = value => value.toLowerCase().replace(/[-_.]+/g, "-");
function search() {
  const query = key(latest.trim());
  if (!query) return;
  const exact = [], prefixes = [], substrings = [];
  let count = 0;
  for (const name of names) {
    const normalized = key(name);
    if (!normalized.includes(query)) continue;
    count++;
    const matches = normalized === query ? exact : normalized.startsWith(query) ? prefixes : substrings;
    if (matches.length < 12) matches.push(name);
  }
  postMessage({ query: latest, results: [...exact, ...prefixes, ...substrings].slice(0, 12), count });
}

onmessage = async ({ data }) => {
  latest = data;
  if (names) return search();
  if (loading) return;
  try {
    loading = load();
    names = await loading;
    search();
  } catch {
    postMessage({ query: latest, error: "Could not load the PyPI index. Check your connection and try again." });
  } finally { loading = null; }
};
