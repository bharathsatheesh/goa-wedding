# Anushri & Bharath · Goa wedding site

A one-page site with personal RSVP links. Hosted free on GitHub Pages. Answers are saved to a Google Sheet.

```
index.html          the whole site. Edit the CONFIG block near the bottom.
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

## Adding guests
In **Guests**, fill in **Name**, and optionally **Household**. Then use **Wedding RSVP → Create links** in the Sheet's menu. That fills in **Code**, **Link** and a ready-to-paste **Message** (on the first row of each household). Codes never change once made, so it's safe to re-run whenever you add people. The **Summary** tab keeps live counts.

## Common edits
- **Event times, RSVP date, wording:** `CONFIG` in `index.html`. Event `id`s must stay in sync with `EVENTS` in `Code.gs`.
- **Photos:** drop them into `img/` (about 1600px wide is plenty) and add a line to `CONFIG.photos`.
- **Ask about food later:** set `askDietary: true`. Guests reuse the same link, and answers land in the **Dietary** column.
- **Changing Code.gs after deploying:** use **Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy**. That keeps the same URL. Creating a *new* deployment gives you a new URL instead.
- **Link preview image/URL:** the `og:` tags at the top of `index.html` need the full site address.
