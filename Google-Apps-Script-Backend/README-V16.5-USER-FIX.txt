V16.5 USER FIX

Fixes in this build:
1. Faster login by removing a blocking Sessions-sheet append from the login path.
2. User Sansad access matching accepts values such as 1, 01, Sansad 1, etc.
3. Non-admin users cannot see or open Settings.
4. Quick Save is device-first: no network refresh is awaited after capture.
5. Photo uploads run in the background with a single sync worker.
6. Activity API is now available.
7. GPS uses a shorter timeout and allows a recent cached fix.
8. Duplicate photo submissions carry a client Photo ID for future server-side idempotency.
