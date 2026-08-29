const CACHE_NAME = "english-memo-v1-1-1";

const FILES_TO_CACHE = [
    "./",
    "./index.html",
    "./manifest.json",
    "./vocabulaire.json",
    "./icon-192.png"
];


/* ==========================================================
   INSTALLATION DU SERVICE WORKER
========================================================== */

self.addEventListener(
    "install",
    function(event) {

        event.waitUntil(

            caches
                .open(CACHE_NAME)
                .then(
                    function(cache) {

                        return cache.addAll(
                            FILES_TO_CACHE
                        );

                    }
                )
                .then(
                    function() {

                        return self.skipWaiting();

                    }
                )

        );

    }
);


/* ==========================================================
   ACTIVATION ET SUPPRESSION DES ANCIENS CACHES
========================================================== */

self.addEventListener(
    "activate",
    function(event) {

        event.waitUntil(

            caches
                .keys()
                .then(
                    function(cacheNames) {

                        return Promise.all(

                            cacheNames.map(
                                function(cacheName) {

                                    if (
                                        cacheName
                                        !== CACHE_NAME
                                    ) {

                                        return caches.delete(
                                            cacheName
                                        );

                                    }

                                    return Promise.resolve();

                                }
                            )

                        );

                    }
                )
                .then(
                    function() {

                        return self.clients.claim();

                    }
                )

        );

    }
);


/* ==========================================================
   GESTION DES REQUÊTES
========================================================== */

self.addEventListener(
    "fetch",
    function(event) {

        if (
            event.request.method
            !== "GET"
        ) {
            return;
        }

        const requestUrl =
            new URL(
                event.request.url
            );


        /*
         * Pour vocabulaire.json :
         *
         * 1. On essaie d'abord la version en ligne.
         * 2. On enregistre la nouvelle version dans le cache.
         * 3. Si Internet ne fonctionne pas, on utilise le cache.
         */

        if (
            requestUrl.pathname.endsWith(
                "/vocabulaire.json"
            )
        ) {

            event.respondWith(

                fetch(
                    event.request
                )
                    .then(
                        function(networkResponse) {

                            if (
                                !networkResponse
                                || networkResponse.status
                                !== 200
                            ) {

                                return networkResponse;

                            }

                            const responseCopy =
                                networkResponse.clone();

                            caches
                                .open(CACHE_NAME)
                                .then(
                                    function(cache) {

                                        cache.put(
                                            "./vocabulaire.json",
                                            responseCopy
                                        );

                                    }
                                );

                            return networkResponse;

                        }
                    )
                    .catch(
                        function() {

                            return caches.match(
                                "./vocabulaire.json"
                            );

                        }
                    )

            );

            return;
        }


        /*
         * Pour les autres fichiers :
         *
         * 1. On regarde d'abord dans le cache.
         * 2. Si le fichier n'est pas présent, on le télécharge.
         * 3. On conserve ensuite une copie dans le cache.
         */

        event.respondWith(

            caches
                .match(
                    event.request
                )
                .then(
                    function(cachedResponse) {

                        if (cachedResponse) {

                            return cachedResponse;

                        }

                        return fetch(
                            event.request
                        )
                            .then(
                                function(networkResponse) {

                                    if (
                                        !networkResponse
                                        || networkResponse.status
                                        !== 200
                                    ) {

                                        return networkResponse;

                                    }

                                    const responseCopy =
                                        networkResponse.clone();

                                    caches
                                        .open(CACHE_NAME)
                                        .then(
                                            function(cache) {

                                                cache.put(
                                                    event.request,
                                                    responseCopy
                                                );

                                            }
                                        );

                                    return networkResponse;

                                }
                            );

                    }
                )

        );

    }
);
