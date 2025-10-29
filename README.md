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
- WhatsApp importer (Gemini-powered) that reads the last 30 days of chat history, extracts action items, and drops them into the "GENERAL -> General Project" board automatically.

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

1. **Netlify / Cloudflare Pages (recommended)** - the repo already contains `netlify.toml`. Set the base directory to `client`, build command `npm run build`, and publish directory `dist`. Add the following environment variables:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
   - `VITE_SHARED_EMAIL`
   - `VITE_SHARED_PASSWORD`
   - `VITE_GEMINI_API_KEY`
   - `VITE_GEMINI_MODEL` (optional, defaults to `gemini-1.5-flash`)
   - `VITE_WHATSAPP_LOOKBACK_DAYS` (optional, defaults to `30`)
   - `VITE_WHATSAPP_COMPANY_NAME` (defaults to `GENERAL`)
   - `VITE_WHATSAPP_PROJECT_NAME` (defaults to `General Project`)
   - `VITE_WHATSAPP_ENDPOINT` (optional, defaults to `v1`)
   - `VITE_WHATSAPP_MAX_LINES` (optional prompt safeguard)
   - `VITE_WHATSAPP_LOG_SHEET_ID` (optional, reserved for upcoming spreadsheet logging)
2. **Manual** - Inside `task-manager/client` run `npm run build` and upload the generated `dist/` folder to any static host.

## WhatsApp Action Item Importer

1. Export the chat from WhatsApp (without media) on mobile. ZIP exports are supported; the importer automatically opens the `.txt` transcript inside the archive.
2. Upload the export via **Import WhatsApp** in the workspace header. Only the last 30 days are analysed, and the importer remembers the latest processed timestamp to avoid duplicates.
3. Pick the Gemini model and endpoint from the dropdown if you need to override the defaults (the values come from the `VITE_*` variables).
4. Google Gemini parses the transcript, extracts action items, respects any explicit or strongly implied deadlines, and only assigns tasks if the named person already exists as a member.
5. Tasks are created in the configured destination (`GENERAL` company -> `General Project`) with the original message timestamp stored as the task's creation date.

> Tip: When you choose a `-latest` or numbered Gemini alias, the importer automatically retries the base model name and the alternate endpoint if the first request is not supported, so you rarely have to guess the exact combination manually.

Environment variables control the behaviour and can be overridden per deployment: `VITE_GEMINI_MODEL`, `VITE_WHATSAPP_LOOKBACK_DAYS`, `VITE_WHATSAPP_COMPANY_NAME`, `VITE_WHATSAPP_PROJECT_NAME`, `VITE_WHATSAPP_ENDPOINT`, `VITE_WHATSAPP_MAX_LINES`, and the optional `VITE_WHATSAPP_LOG_SHEET_ID` for future spreadsheet logging.

> Media OCR and spreadsheet logging hooks are stubbed in the codebase and can be enabled later without changing the importer UI.
## Customisation Notes

- Adjust colours, typography, or spacing near the top of `styles.css`.
- Extend the task model in `src/app.js` (see `addTask` and `updateTask`) if you need more fields or workflow rules.
- Seed default sections, departments, or members by tweaking the constants near the top of `src/app.js`.
- Tailwind via CDN powers the UI. Swap it with a PostCSS build if you prefer local Tailwind compilation.
- The client signs into Supabase automatically using the shared credentials in `VITE_SHARED_EMAIL` / `VITE_SHARED_PASSWORD`. Create that user once in Supabase Auth (email/password) and share the same values with your team.
- The WhatsApp importer needs Gemini credentials and a destination company/project (see `VITE_*` variables above).

## Supabase Migration Status

- `supabase/schema.sql` defines workspaces, projects, task, attachment, and team tables plus RLS policies.
- Storage bucket `attachments` should be created in Supabase with read/insert policies limited to authenticated users.
- Supabase client bootstrap uses `VITE_SUPABASE_URL` + `VITE_SUPABASE_ANON_KEY`; keep the service role key private for server-side work.
- Dashboard and settings auto-connect with the shared credentials. Projects, sections, tasks, members, and departments now read/write from Supabase. Settings and realtime updates are the remaining items to migrate.

## Browser Support

Tested on current Chrome, Edge, Safari, and Firefox. When `crypto.randomUUID` is unavailable the app falls back to timestamp-based identifiers automatically.

## License

MIT - free for personal or commercial use. Update this section if you ship with different terms.
