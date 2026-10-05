const CACHE_NAME =
    "verbos-irregulares-v3";

/* ==========================
   ARCHIVOS A CACHEAR
========================== */

const urlsToCache = [

    "./",

    "./index.html",

    "./style.css",

    "./app.js",

    "./manifest.json",

    "./xlsx.full.min.js",

    "./data/Verbos_ingles.xlsx",

    "./icons/icon-192.png",

    "./icons/icon-512.png"

];

/* ==========================
   INSTALACIÓN
========================== */

self.addEventListener(
    "install",
    event => {

        event.waitUntil(

            caches.open(
                CACHE_NAME
            )

            .then(cache => {

                console.log(
                    "Cache creada"
                );

                return cache.addAll(
                    urlsToCache
                );

            })

        );

    }
);

/* ==========================
   ACTIVACIÓN
========================== */

self.addEventListener(
    "activate",
    event => {

        event.waitUntil(

            caches.keys()

            .then(keys => {

                return Promise.all(

                    keys.map(key => {

                        if (
                            key !==
                            CACHE_NAME
                        ) {

                            return caches.delete(
                                key
                            );

                        }

                    })

                );

            })

        );

    }
);

/* ==========================
   FETCH
========================== */

self.addEventListener(
    "fetch",
    event => {

        event.respondWith(

            caches.match(
                event.request
            )

            .then(response => {

                if (response) {

                    return response;

                }

                return fetch(
                    event.request
                );

            })

        );

    }
);