# Astra initial review, 2026-10-10

Model: gpt-6-astra, xhigh. Frozen source snapshot. Verdict: FIX-FIRST.
Proofs used synthetic in-memory SQLite and fake transports, with no active data mutation.

1. HIGH: Engine deltas erase full pp_results snapshots. formualizer.py:142-151,
   calc_service.py:23-28. Empty revision2 delta leaves results empty while current.
2. HIGH: Fork revision reconstruction reads future inherited values. ops.py:118-123,
   85-87. Fork rev1 A12,B7 becomes A99,B7 after rev2 changes A.
3. HIGH: category-to-private PATCH does not revoke existing sockets. server.py:154-157,
   live.py:37-44. Old category member receives subsequent private op.
4. HIGH: existing sockets survive logout/expiry. server.py:103-109, live.py:95-96,
   113-116. Logged-out socket hello still returns calc_state.
5. HIGH: concurrent broadcasts can deliver calc revisions backwards. live.py:37-44,
   server.py:220-222. Delayed rev1 send permits recipients to see [2,1].
6. HIGH: reconnect last_revision=1 omits calc1 completed during disconnect.
   live.py:67-70. Only calc_state(1,1) sent, leaving displayed values stale.

Repair full worker snapshots, derive event deltas against the last published DB
snapshot, immutable fork baseline, session/access checks on live traffic, ordered
per-copy publication and a safe same-revision reconnect resnapshot/replay.

Known root findings: fresh revision0 formulas blank despite current marker,
shared ISFORMULA followers not rewritten, bool/number comparisons conflated.
Full engine and test suite were not run by this review; controller owns those checks.

## Follow-up review, 2026-10-10

Original six HIGH defects resolved. Independent32 synthetic tests passed; frozen manifest74 files unchanged. Verdict FIX-FIRST on three new HIGHs: API error-vs-text equality, reuse of dead copy workers, and actor reauthorization stranding accepted calculation after revocation. Four MEDIUMs: baseline supersession during creation, malformed WebSocket traffic, unsupported error tokens accepted before calculation, and case-sensitive ISFORMULA references. All seven repairs are assigned with test-first regressions. Final follow-up pending.
