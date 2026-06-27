# Personal Website

A modern, fast, fully responsive personal-brand website built with **Next.js (App Router)**, **TypeScript**, **Tailwind CSS v4**, and **Framer Motion**.

Inspired by the structure of personal-brand sites like ankurwarikoo.com — hero statement, "as featured on", what-I-do roles, ventures, selected work, social/connect, writing, and a contact form.

## Customize your site (the only file you need)

All content lives in **`src/content/site.ts`**. Edit that one file to make the site yours:

- `meta` — your name, role, page title/description (SEO), site URL, and logo initials
- `hero` — your headline, taglines, intro, call-to-action buttons, and photo
- `featuredOn` — press/media names shown in the marquee
- `about.roles` — the "What I do" cards
- `connect` — follower stats and social links
- `ventures` — companies/projects you're building
- `work` — books / products / talks
- `writing.posts` — recent blog posts
- `contact` — your email and (optional) form endpoint

### Add your photo
Drop an image into the `public/` folder (e.g. `public/me.jpg`) and set `hero.photo: "/me.jpg"` in `site.ts`.

### Make the contact form send real emails
By default the form opens the visitor's email client (mailto). To collect submissions instead, create a free form endpoint (e.g. [Formspree](https://formspree.io) or [Getform](https://getform.io)) and paste the URL into `contact.formEndpoint` in `site.ts`.

### Change the colors / theme
Edit the design tokens at the top of `src/app/globals.css` (`--accent`, `--background`, etc.). Light and dark mode are both supported automatically.

## Run locally

```bash
npm install
npm run dev      # http://localhost:3000
```

## Build for production

```bash
npm run build
npm run start
```

## Deploy

This is a standard Next.js app and deploys for free on:
- **Vercel** — import the repo at [vercel.com/new](https://vercel.com/new)
- **Netlify** — connect the repo; build command `npm run build`

## Project structure

```
src/
  app/
    layout.tsx        # fonts + SEO metadata
    page.tsx          # composes all sections
    globals.css       # design tokens + theme
  components/         # one component per section
    ui/               # Section, Reveal (animation), Icon helpers
  content/
    site.ts           # ← EDIT THIS to personalize everything
```
