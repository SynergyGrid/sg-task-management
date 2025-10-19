# Synergy Tasks

Synergy Tasks is a Todoist-style task manager that runs completely in the browser. Inbox, Today, Upcoming, and project views mirror familiar productivity flows while everything stays on the device via `localStorage`.

## Highlights

- Inbox, Today, Upcoming views with automatic counts and overdue badges.
- Sidebar projects with quick add, colour accents, per-project task totals, and custom sections.
- List <-> Board switcher for each project, including drag-and-drop across sections.
- Team directory with departments and members that can be assigned to tasks.
- Quick-add form for new tasks plus a full editor dialog for deeper changes.
- Priorities, due dates, descriptions, and completion history with a toggle.
- Type-to-filter search, JSON export, and a full data reset option.

## Folder Layout

```
task-manager/
├─ client/                 # Vite application
│  ├─ index.html           # Workspace dashboard entry
│  ├─ settings.html        # Profile & theme editor
│  ├─ src/
│  │  ├─ app.js            # Workspace logic (currently localStorage-backed)
│  │  ├─ settings.js       # Settings page logic
│  │  ├─ styles.css        # Tailwind-enhanced styling
│  │  └─ lib/
│  │     └─ supabaseClient.js # Supabase client bootstrap
│  └─ vite.config.js       # Multi-page build configuration
├─ supabase/
│  └─ schema.sql           # Database schema + RLS policies
├─ .env.example            # Example env file for Supabase creds
└─ .env.local              # Local-only env values (gitignored)
```

From `task-manager/client` run:

```
npm install
npm run dev
```

The dev server runs at `http://localhost:5173` by default and serves both `index.html` and `settings.html`.

## Deploy Online

The app is static, so any static host works. Two quick paths:

1. **Netlify** - connect the repo, set the base directory to `task-manager/client`, build command `npm run build`, and publish directory `task-manager/client/dist`.
2. **Manual** - run `npm run build` inside `task-manager/client` and upload the generated `dist/` folder to any static host.

## Customisation Notes

- Adjust colours, typography, or spacing near the top of `styles.css`.
- Extend the task model in `app.js` (see `addTask` and `updateTask`) if you need more fields or workflow rules.
- Seed default sections, departments, or members by tweaking the constants near the top of `app.js`.
- Swap `loadJSON`/`saveJSON` for your own API calls to support shared workspaces.
- Tailwind via CDN powers the new UI; adjust the utility classes in `index.html` or replace the CDN link with your own build as needed.

## Supabase Migration

- `supabase/schema.sql` defines workspaces, projects, tasks, attachments, team directory tables, and row-level security policies aligned with the Supabase rollout plan.
- Run the script in a fresh Supabase project before wiring the frontend; it enables RLS so use the service role key for migrations and provide the anon key to the client.
- Create a storage bucket called `attachments` and restrict its access to workspace members once the Supabase dashboard is ready.
- Populate Netlify (or Vite) environment variables with `VITE_SUPABASE_URL` and `VITE_SUPABASE_ANON_KEY`; keep the service role key secret for server-side tasks only.

## Browser Support

Tested on current Chrome, Edge, Safari, and Firefox. When `crypto.randomUUID` is unavailable the app falls back to timestamp-based identifiers automatically.

## License

MIT - free for personal or commercial use. Update this section if you ship with different terms.
