# LinkedIn ad references: Fluso, Prem Enclave, Enclave API

Source for the Google Doc "LinkedIn Ad References — Fluso · Prem Enclave · Enclave API (Oct 2026)".

- `concepts.json`: the 18 concepts (3 static + 3 video per product), each with its reference ad, copy, storyboard and targeting.
- `build_doc.py`: renders `concepts.json` into `linkedin-creative-references.html`, which is imported into Google Docs through Drive (HTML upload converts to a Doc).
- Reference images are YouTube thumbnails of the source ads. Concept mockups are Higgsfield generations (gpt_image_2_5) hosted on Higgsfield's CDN.

Rebuild with `python3 build_doc.py`.
