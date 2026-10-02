# Performance notes v9

- Google Apps Script now caches Events and Participants for 30 seconds and per-event attendance for 5 seconds.
- Write operations invalidate the relevant caches.
- Public name search debounce increased to 700 ms to reduce repeated backend calls while typing.
- Admin dashboard no longer performs the duplicate initial `adminLoad` request caused by selecting the first event.
- No behavior or data schema changes are intended.
