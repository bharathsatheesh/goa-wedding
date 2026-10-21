# Anushri & Bharath · Goa wedding site

A one-page site with personal RSVP links. Hosted free on GitHub Pages at https://bharathsatheesh.github.io/goa-wedding/ (served from the `gh-pages` branch, so push changes to both `main` and `gh-pages`). Answers are saved to a Google Sheet.

```
index.html          the site ("spin the mandala"). It is generated: don't edit it by hand.
src/content.html    the content and RSVP code (wording, CONFIG, events, photos, form). Edit here,
                    then rebuild index.html with the build script (build-spin.py + spin-template.html).
spin/index.html     forwards the old preview address to the main page, keeping ?g= codes.
img/                photos, card artwork, og.jpg (the WhatsApp/iMessage link preview)
apps-script/Code.gs the Sheet backend (it gets pasted into the Sheet, not served from here)
```

## How the links work
Each guest row in the Sheet gets a 6-character code. Their link is `SITE/?g=CODE`. When the page opens, it asks the Sheet which names belong to that code and shows only those names. Nobody can add a guest. People who share a **Household** value share one link and answer for each person on it.

The repo is public, so guest names live only in the Sheet and never in this repo.

## One-time Google setup (about 5 min)
1. Open the RSVP Sheet. Go to **Extensions → Apps Script**, delete the sample code, paste in `apps-script/Code.gs`, and save.
2. Pick `setup` in the function dropdown and click **Run**. Approve the permissions prompt. (Google warns that the app is unverified because it's your own script: **Advanced → Go to … (unsafe)**.) This creates the **Guests** and **Summary** tabs, with a test row for each of you.
3. Click **Deploy → New deployment → Web app**. Set *Execute as*: **Me** and *Who has access*: **Anyone**, then click **Deploy** and copy the URL that ends in `/exec`.
   - If **Anyone** isn't offered, your organisation's account blocks it. Make a copy of the Sheet in a personal Gmail account (**File → Make a copy**) and repeat from step 1 there.
4. Paste that URL into `CONFIG.apiUrl` in `index.html`.

## Editing the guest list (no code needed)
Everything lives in the **Guests** tab. The site reads it live, and the Sheet keeps itself tidy as you type:
- **Rename someone:** edit their **Name**. The RSVP form and their household's invite message update.
- **Add someone:** type their **Name** and **Household**. They get a code, link and message straight away. Use the same Household as someone else to share their link.
- **Split a pair:** change one person's **Household** (e.g. "Anushri's friends · Marissa"). They get their own link and message, and the other person's message is updated.
- **Remove someone:** delete their row.

Messages someone has edited by hand are never overwritten (a notice pops up instead); clear the cell to get a fresh one. Households starting "Anushri's friends" or "Bharath's friends" get the friends wording; everyone else gets the family wording. **Wedding RSVP → Create links** fills in anything missing after a big paste. The **Summary** tab keeps live counts.

## Common edits
- **Event times, RSVP date, wording:** `CONFIG` in `src/content.html`, then rebuild `index.html`. Event `id`s must stay in sync with `EVENTS` in `Code.gs`.
- **Photos:** drop them into `img/` (about 1600px wide is plenty) and add a line to `CONFIG.photos`.
- **Ask about food later:** set `askDietary: true`. Guests reuse the same link, and answers land in the **Dietary** column.
- **Changing Code.gs after deploying:** use **Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy**. That keeps the same URL. Creating a *new* deployment gives you a new URL instead.
- **Link preview image/URL:** the `og:` tags at the top of `index.html` need the full site address.
