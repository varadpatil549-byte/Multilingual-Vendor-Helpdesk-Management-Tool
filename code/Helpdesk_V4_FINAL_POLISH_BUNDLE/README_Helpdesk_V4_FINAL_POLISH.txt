AXALTA MULTILINGUAL VENDOR HELPDESK — V4 FINAL POLISH

1. Enlarged/redesigned application title.
2. Fixed Outlook ReceivedTime display; locale-specific COM strings are parsed
   instead of taking the last five characters, which caused 00:00.
3. Query translation runs in a background worker so the Qt window remains
   responsive while Ollama is processing.
4. Response translation also runs in a background worker.
5. Stale worker results are ignored after switching emails/manual sessions.
6. Existing cache, latest-message-only extraction, priority triage, integrity
   percentage and manual-send workflow are retained.

Validation: AST parse + Python compile() PASS.
