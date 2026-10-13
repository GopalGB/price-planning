# Backend security verification

Development dependency audit used the complete hashed requirements-dev.lock,
including runtime requirements. Result: no known vulnerabilities. Starlette was
updated to 1.3.1 after the earlier pinned version failed this audit; both lockfiles
and Windows wheel inventory were regenerated together.

Semgrep p/python and p/owasp-top-ten initially returned zero findings and zero
errors across pp/, run_v2.py and ops/lock_deps.py. Final repair scan is pending.

Initial Bandit scan returned 18 findings, none HIGH. Its XML warnings concern
configured local templates; this API has no arbitrary file-upload endpoint.
The dynamic SQL history filters use fixed identifiers and parameterized values.
The remaining LOW notices concern fixed subprocess commands and disconnected
socket cleanup. Final scan and reviewer assessment remain pending.

Gitleaks staged scans passed before the existing backend commits. Final staged
scan is required before the remaining commits. Paid Strix, external penetration
testing and Office/LAN security acceptance were not run or claimed.
