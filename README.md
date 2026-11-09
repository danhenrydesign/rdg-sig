# RDG Signatures (Outlook add-in)

Adds a signature that matches the From address:
- `@revealdesigngroup.com`: Reveal Design Group text
- `@coda-lp.com`: Coda text
- any other domain: no signature

The formats match `example-signatures/`. New mail gets the full format, with address and links. Replies and forwards get the short one.
Name comes from the user's account. The From address decides the brand.
Each user sets their pronouns, title and phone extension once, using **Signature settings** on the compose ribbon.
These values are stored in the user's mailbox (roaming settings).

Edit the company text, phone numbers and links in `BRANDS`, and the layout in `buildSignature`. Both are in `src/commands.js`. Then run `node test.js`.

## 1. Host the files (GitHub Pages)
1. Create a GitHub repo and push this folder to it. Free Pages needs a public repo. Nothing in here is secret.
2. In the repo, go to Settings → Pages → Deploy from branch `main`, folder `/ (root)`.
3. Your base URL is `https://<github-user>.github.io/<repo>`.
4. In `manifest.xml`, replace every `https://YOURHOST` with that base URL.
5. Validate: `npx office-addin-manifest validate manifest.xml`.

Any static HTTPS host works, for example Azure Static Web Apps.
Changes to the JS or HTML go live when you push. Changes to the manifest need re-uploading in step 3.

## 2. Test on your own machine
In classic Outlook, go to Get Add-ins → My add-ins → Custom add-ins → Add from file → `manifest.xml`.
Then open a new message, then a reply, then switch From to the coda-lp.com alias.

## 3. Roll out to users
Go to the M365 admin center → Settings → Integrated apps → Upload custom apps → Office Add-in → upload `manifest.xml`, then assign it to users or groups.
It can take up to 24 hours to appear. Users need classic Outlook version 2304 (build 16327) or later.

Tell users to set their own Outlook signature to (none). The add-in also turns it off for each message.

## Notes
- `assets/*.png` are plain placeholder squares. Replace them with real icons that keep the same sizes.
