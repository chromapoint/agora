# Agora
### Where your story finds its pulse.

Turns a HartChart export into the two documents Story Compass needs:
a Story Treatment and the Core & Bonus Questions.

## Deploy to Vercel (5 minutes)

1. **Get an Anthropic API key** at console.anthropic.com → API Keys → Create Key.
   Add a few dollars of credit under Billing.

2. **Sign up at vercel.com** (free).

3. **Add New → Project**, drag this unzipped folder into the import area.

4. Before deploying, open **Environment Variables** and add:
   - Name:  `ANTHROPIC_API_KEY`
   - Value: your `sk-ant-…` key

5. **Deploy.** You'll get a live URL in ~60 seconds.

## Local development

```
npm install
npm install -g vercel
vercel dev        # runs at localhost:3000
```

## Structure

- `public/index.html` — the Agora web app (React via CDN)
- `api/claude.js`     — serverless relay; holds the API key, two modes:
                        passthrough (chat) and generate (builds .docx)
- `api/pipeline.js`   — extract → write → build both documents
- `api/generate.js`   — standalone document generator
- `vercel.json`       — routing config

## Changing colours & fonts

Open `public/index.html`. The theme is at the top:
- CSS `:root` block — colours
- the Google Fonts `<link>` — swap fonts here
- the `THEME` object in the script — JS-side tokens
