# SPX Vendor Price Portal

Public price-submission page for external manpower vendors (GitHub Pages) + Google Apps Script code for the
vendor submission sheet and the Manpower Bidding Price Tracker.

- **Vendors** open the GitHub Pages link, sign in with company name + access code, submit prices per facility.
- **Prices** are saved through a hidden Google Form into a private company Google Sheet.
- **The tracker** pulls them daily at 18:00 (Bangkok) or on demand, and updates the Tracker and Trend tabs.

Working with Claude Code: open this folder in Claude Code — it reads `CLAUDE.md` for the architecture, rules and
step-by-step tasks. Typical requests:

- "Set this repo up on GitHub and publish the page" (task A)
- "Here is the new config.js from Drive: ~/Downloads/config.js — publish it" (task B)
- "Change the page heading / add a hint under Max HC" (task C)
- "Add a required field 'Contract months'" (task D)

Checks: `npm test` (Node 20+). CI runs the same tests and only deploys `docs/` when they pass.
