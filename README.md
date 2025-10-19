# Synergy Tasks

Synergy Tasks is a Todoist-style task manager. The UI now boots directly into a shared Supabase workspace (no login screen), so every browser using the shared credentials sees the same projects, sections, tasks, members, and departments.

## Highlights

- Inbox, Today, Upcoming views with automatic counts and overdue badges.
- Sidebar projects with quick add, colour accents, per-project task totals, and custom sections.
- List ? Board switcher for each project, including drag-and-drop across sections.
- Team directory with departments and members that can be assigned to tasks.
- Quick-add form for new tasks plus a full editor dialog for deeper changes.
- Priorities, due dates, descriptions, and completion history with a toggle.
- Type-to-filter search, JSON export, and full data reset options.

## Folder Layout

```
task-manager/
|- client/                 # Vite application
|  |- index.html           # Workspace dashboard entry
|  |- settings.html        # Theme/profile editor
|  |- src/
|  |  |- app.js            # Workspace logic (Supabase + local cache)
|  |  |- settings.js       # Settings page logic
|  |  |- styles.css        # Tailwind-enhanced styling
|  |  |- main.js           # Auto-connects to shared Supabase session
|  |  |- settings.main.js  # Loads settings once session established
|  |  |- lib/
|  |     |- supabaseClient.js  # Supabase client bootstrap
|  |     |- sharedSession.js   # Shared credential session helper
|  |- vite.config.js       # Multi-page build configuration
|- supabase/
|  |- schema.sql           # Database schema + RLS policies
|- .env.example            # Example env file (Supabase + shared login)
|- netlify.toml            # Netlify build configuration
```

## Local Development

From `task-manager/client`:

```
npm install
npm run dev
```

The dev server runs at `http://localhost:5173` and serves `index.html` and `settings.html`. On start-up the client signs into Supabase automatically using the shared credentials (see environment variables below).

## Deploy Online

1. **Netlify / Cloudflare Pages (recommended)** – the repo already contains `netlify.toml`. Set the base directory to `client`, build command `npm run build`, and publish directory `dist`. Add the following environment variables:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
   - `VITE_SHARED_EMAIL`
   - `VITE_SHARED_PASSWORD`
2. **Manual** – Inside `task-manager/client` run `npm run build` and upload the generated `dist/` folder to any static host.

## Customisation Notes

- Adjust colours, typography, or spacing near the top of `styles.css`.
- Extend the task model in `src/app.js` (see `addTask` and `updateTask`) if you need more fields or workflow rules.
- Seed default sections, departments, or members by tweaking the constants near the top of `src/app.js`.
- Tailwind via CDN powers the UI. Swap it with a PostCSS build if you prefer local Tailwind compilation.
- The client signs into Supabase automatically using the shared credentials in `VITE_SHARED_EMAIL` / `VITE_SHARED_PASSWORD`. Create that user once in Supabase Auth (email/password) and share the same values with your team.

## Supabase Migration Status

- `supabase/schema.sql` defines workspaces, projects, task, attachment, and team tables plus RLS policies.
- Storage bucket `attachments` should be created in Supabase with read/insert policies limited to authenticated users.
- Supabase client bootstrap uses `VITE_SUPABASE_URL` + `VITE_SUPABASE_ANON_KEY`; keep the service role key private for server-side work.
- Dashboard and settings auto-connect with the shared credentials. Projects, sections, tasks, members, and departments now read/write from Supabase. Settings and realtime updates are the remaining items to migrate.

## Browser Support

Tested on current Chrome, Edge, Safari, and Firefox. When `crypto.randomUUID` is unavailable the app falls back to timestamp-based identifiers automatically.

## License

MIT – free for personal or commercial use. Update this section if you ship with different terms.
