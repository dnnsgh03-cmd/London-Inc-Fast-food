const CACHE_NAME = "london-inc-shell-v7";
const APP_SHELL = [
    "/static/style.css",
    "/static/script.js",
    "/static/manifest.json",
    "/static/img/logo.jpg",
    "/static/img/pwa-icon.svg",
];

self.addEventListener("install", event => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => cache.addAll(APP_SHELL))
            .then(() => self.skipWaiting())
    );
});

self.addEventListener("activate", event => {
    event.waitUntil(
        caches.keys()
            .then(keys => Promise.all(
                keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key))
            ))
            .then(() => self.clients.claim())
    );
});

self.addEventListener("fetch", event => {
    const requestUrl = new URL(event.request.url);
    if (
        event.request.method !== "GET"
        || requestUrl.origin !== self.location.origin
        || requestUrl.pathname.startsWith("/api/")
        || !requestUrl.pathname.startsWith("/static/")
    ) return;

    event.respondWith(
        fetch(event.request)
            .then(response => {
                if (!response.ok) return response;
                return caches.open(CACHE_NAME)
                    .then(cache => cache.put(event.request, response.clone()))
                    .then(() => response);
            })
            .catch(error => caches.match(event.request, { ignoreSearch: true })
                .then(cached => {
                    if (cached) return cached;
                    throw error;
                }))
    );
});
