# Synergy Tasks

Synergy Tasks is a Todoist-style task manager. The current UI is still localStorage-powered, but the project is now scaffolded to plug into Supabase for authentication, shared data, and realtime sync.

## Highlights

- Inbox, Today, Upcoming views with automatic counts and overdue badges.
- Sidebar projects with quick add, colour accents, per-project task totals, and custom sections.
- List ↔ Board switcher for each project, including drag-and-drop across sections.
- Team directory with departments and members that can be assigned to tasks.
- Quick-add form for new tasks plus a full editor dialog for deeper changes.
- Priorities, due dates, descriptions, and completion history with a toggle.
- Type-to-filter search, JSON export, and full data reset options.

## Folder Layout

```
task-manager/
|- client/                 # Vite application
|  |- index.html           # Workspace dashboard entry
|  |- settings.html        # Profile & theme editor
|  |- auth.html            # Sign-in/Sign-up screen (Supabase auth)
|  |- src/
|  |  |- app.js            # Workspace logic (still localStorage-backed)
|  |  |- settings.js       # Settings page logic
|  |  |- styles.css        # Tailwind-enhanced styling
|  |  |- auth.main.js      # Auth screen controller
|  |  |- main.js           # Dashboard entry (guards session)
|  |  |- settings.main.js  # Settings entry (guards session)
|  |  |- lib/
|  |     |- supabaseClient.js # Supabase client bootstrap
|  |     |- authGuard.js   # Session gate shared by pages
|  |- vite.config.js       # Multi-page build configuration
|- supabase/
|  |- schema.sql           # Database schema + RLS policies
|- .env.example            # Example env file for Supabase creds
|- .env.local              # Local-only env values (gitignored)
|- netlify.toml            # Netlify build configuration
```

## Local Development

From `task-manager/client`:

```
npm install
npm run dev
```

The dev server runs at `http://localhost:5173` and serves `index.html`, `settings.html`, and `auth.html`. Visit `/auth.html` to create a Supabase user or sign in; the dashboard and settings pages redirect there automatically if no session is present.

## Deploy Online

1. **Netlify (recommended)** – Netlify reads `netlify.toml` and uses base `client`, build `npm run build`, publish `dist`. Add `VITE_SUPABASE_URL` and `VITE_SUPABASE_ANON_KEY` in Site → Environment variables before deploying.
2. **Manual** – Inside `task-manager/client` run `npm run build` and upload the generated `dist/` folder to any static host.

## Customisation Notes

- Adjust colours, typography, or spacing near the top of `styles.css`.
- Extend the task model in `src/app.js` (see `addTask` and `updateTask`) if you need more fields or workflow rules.
- Seed default sections, departments, or members by tweaking the constants near the top of `src/app.js`.
- Tailwind via CDN powers the UI. Swap it with a PostCSS build if you prefer local Tailwind compilation.
- `auth.html` currently supports email/password sign-up & sign-in. Additional providers (magic link, OAuth) can be added via Supabase Auth.

## Supabase Migration Status

- `supabase/schema.sql` defines workspaces, projects, task, attachment, and team tables plus RLS policies.
- Storage bucket `attachments` should be created in Supabase with read/insert policies limited to authenticated users.
- Supabase client bootstrap uses `VITE_SUPABASE_URL` + `VITE_SUPABASE_ANON_KEY`; keep the service role key private for server-side work.
- Dashboard and settings now require a Supabase session before loading. Next steps are swapping the localStorage persistence for Supabase queries and enabling realtime subscriptions.

## Browser Support

Tested on current Chrome, Edge, Safari, and Firefox. When `crypto.randomUUID` is unavailable the app falls back to timestamp-based identifiers automatically.

## License

MIT – free for personal or commercial use. Update this section if you ship with different terms.
