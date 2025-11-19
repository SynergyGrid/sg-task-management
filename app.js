import JSZip from 'jszip';
import { db } from './lib/firebaseClient';
import { doc, getDoc, onSnapshot, setDoc } from 'firebase/firestore';
const STORAGE_KEYS = {
  tasks: "synergygrid.todoist.tasks.v2",
  projects: "synergygrid.todoist.projects.v2",
  sections: "synergygrid.todoist.sections.v1",
  companies: "synergygrid.todoist.companies.v1",
  members: "synergygrid.todoist.members.v1",
  departments: "synergygrid.todoist.departments.v1",
  preferences: "synergygrid.todoist.preferences.v2",
  settings: "synergygrid.todoist.settings.v1",
  userguide: "synergygrid.todoist.userguide.v1",
  imports: "synergygrid.todoist.imports.v1",
};

const OPENROUTER_MODEL = import.meta.env.VITE_OPENROUTER_MODEL || "openai/gpt-4o-mini";
const OPENROUTER_API_KEY = import.meta.env.VITE_OPENROUTER_API_KEY || "";
const WHATSAPP_LOOKBACK_DAYS = Number.parseInt(import.meta.env.VITE_WHATSAPP_LOOKBACK_DAYS ?? "30", 10);
const WHATSAPP_COMPANY_NAME = import.meta.env.VITE_WHATSAPP_COMPANY_NAME || "";
const WHATSAPP_PROJECT_NAME = import.meta.env.VITE_WHATSAPP_PROJECT_NAME || "WhatsApp Tasks";
const WHATSAPP_SECTION_NAME = import.meta.env.VITE_WHATSAPP_SECTION_NAME || "WhatsApp Tasks";
const WHATSAPP_DEFAULT_PROVIDER = "openrouter";
const MAX_WHATSAPP_LINES = Number.parseInt(import.meta.env.VITE_WHATSAPP_MAX_LINES ?? "2000", 10);
const WHATSAPP_LOG_SHEET_ID = import.meta.env.VITE_WHATSAPP_LOG_SHEET_ID || "";


const DEFAULT_COMPANY = {
  id: "company-default",
  name: "Synergy Grid",
  isDefault: true,
  createdAt: new Date().toISOString(),
};

const ALL_COMPANY_ID = "company-all";

const DEFAULT_PROJECT = {
  id: "inbox",
  name: "Inbox",
  companyId: DEFAULT_COMPANY.id,
  color: "#e44232",
  isDefault: true,
  createdAt: new Date().toISOString(),
};

const DEFAULT_DEPARTMENT = {
  id: "department-general",
  name: "General",
  isDefault: true,
  createdAt: new Date().toISOString(),
};

const DEFAULT_SECTION_NAME = "General";
const PROJECT_OVERVIEW_GROUPS = [
  {
    id: "regular",
    title: "Regular Tasks",
    filter: (task) => (task.kind ?? "task") === "task" && (task.source ?? "manual") !== "whatsapp",
  },
  {
    id: "whatsapp",
    title: "WhatsApp Tasks",
    filter: (task) => (task.source ?? "") === "whatsapp",
  },
  {
    id: "meeting",
    title: "Meetings",
    filter: (task) => (task.kind ?? "task") === "meeting",
  },
  {
    id: "email",
    title: "Emails",
    filter: (task) => (task.kind ?? "task") === "email",
  },
];
const MEETING_TYPES = ["Google Meet", "Teams", "Zoom", "In Person"];
const PROJECT_COLORS = ["#246fe0", "#fa968f", "#ffc247", "#8f3ffc", "#24a564", "#e25dd2", "#00a896"];
const PRIORITY_WEIGHT = { critical: 0, "very-high": 1, high: 2, medium: 3, low: 4, optional: 5 };
const PRIORITY_LABELS = {
  critical: "Critical priority",
  "very-high": "Very high priority",
  high: "High priority",
  medium: "Medium priority",
  low: "Low priority",
  optional: "Optional priority",
};
const MIN_TEXTAREA_HEIGHT = 72;
const MAX_TEXTAREA_HEIGHT = 720;
const THREAD_MESSAGE_MIN_HEIGHT = 40;
const THREAD_MESSAGE_MAX_HEIGHT = 640;
const FILTER_MEMBER_UNASSIGNED = "__unassigned";
const AUTOSIZE_RESET_KEYS = new Set([
  "quick-add-description",
  "dialog-description",
  "meeting-notes",
  "email-notes",
]);
const DEFAULT_USERGUIDE = [
  "## Workspace overview",
  "### Tabs and scope",
  "- Workspace lists unscheduled tasks for the active company.",
  "- Due Today and Upcoming 7 Days highlight urgent work and update their counters automatically.",
  "- Switch companies from the pill strip; use the three-dot menu on each pill to rename or archive a company safely.",
  "### Projects",
  "- Pick a project from the dropdown to review its overview cards (Regular Tasks, WhatsApp Tasks, Meetings, Emails).",
  "- Use the Meeting and Email quick actions beside a project to open the respective creation dialogs instantly.",
  "- Select See more inside any card to expand beyond the latest five records.",
  "## Capturing work",
  "### Tasks",
  "- Click Add Task for a quick capture or open any task row to edit the full detail dialog.",
  "- Description fields start at three lines and expand automatically as you type.",
  "- Task cards surface the first three lines of notes plus assignee, action items, and due dates at a glance.",
  "- Use the Mark completed / Restore control inside the task dialog to switch status.",

  "- Closing the task window, clicking outside, or pressing Save all capture your changes automatically.",
  "### Meetings",
  "- Use the Meeting quick action to log attendees, meeting type, links, notes and follow-up items.",
  "- Action items appear as an interactive checklist. Paste multiple lines to create a list automatically and tick entries off as work lands.",
  "- Meetings surface under the project overview so you can monitor follow-ups without leaving the overview tab.",

  "- Mark meetings complete from the dialog header and any changes are auto-saved when you close.",
  "### Emails",
  "- The Email quick action tracks important threads with status, notes and supporting links.",
  "- Notes resize with the content and remember the height you prefer for long summaries.",

  "- Mark emails complete straight from the header; closing the window auto-saves edits too.",
  "## Finding context",
  "- The global search provides live suggestions; selecting a result jumps to the correct company/project and opens the task dialog.",
  "- All Activity now includes Created, Updated, Completed and Deleted tabs for a richer audit trail.",
  "- Click anywhere on a task (outside of buttons) to open its detail dialog, no need to hunt for the Edit button.",
  "## Administration and exports",
  "- Manage members and departments from Settings so assignments and filters stay accurate across the workspace.",
  "- Use the Workspace export button in Settings to download tasks, projects, sections, companies, members, departments and the userguide as JSON.",
  "- Deleting a project or company preserves its tasks by moving them into Workspace under the default company.",
  "- Members and departments can be edited inline from Settings - adjust names or move people between departments without leaving the page.",
  "## Recent improvements",
  "- Project settings open in a dedicated dialog so you can rename, reassign, or delete with confirmation.",
  "- Meeting and Email quick actions sit on their own row in the project picker for easier tapping.",
  "- Completed tasks highlight a Restore button in the editor, and search jumps now spotlight the matching card.",
  "## Maintaining this guide",
  "- Update the userguide whenever processes change. Headings (`##`, `###`), bullet lists (`- item`) and numbered steps (`1.`) are supported.",
  "- Keep entries concise but complete so teammates can follow the latest workflow without additional training.",
];

const USERGUIDE_LATEST_ENTRIES = [

  "- Task cards surface the first three lines of notes plus assignee, action items, and due dates at a glance.",

  "- Use the Mark completed / Restore control inside the task dialog to switch status.",

  "- Closing any task, meeting, or email dialog now auto-saves your edits.",

  "- Meeting and Email editors include Mark completed toggles so you can finish work without leaving the dialog.",

  "- Settings shows members and departments as editable inline lists for quick updates.",

  "- Every task type can be dragged between sections on the board view.",

  "## Recent improvements",

  "- Project settings open in a dedicated dialog so you can rename, reassign, or delete with confirmation.",

  "- Meeting and Email quick actions sit on their own row in the project picker for easier tapping.",

  "- Completed tasks highlight a Restore button in the editor, and search jumps now spotlight the matching card.",

];

const LEGACY_USERGUIDE_V1 = [
  "Pick a company. Use the Company dropdown to jump between organisations. Each one keeps its own projects and sections.",
  "Choose a project. The Project dropdown lists only work that belongs to the selected company and remembers the last project you touched for quick access.",
  "Plan with sections. Sections stay visible on the board. Drag them to reorder and use the menu in each column to rename or remove them.",
  "Add or move tasks. Create tasks from Quick Add or the board, then drag them between sections or edit them to move across companies.",
  "Keep the workspace tidy. Archive projects or companies you no longer need, and revisit this guide whenever new features roll out.",
];

const LEGACY_USERGUIDE_V2 = [
  "Choose a company from the dropdown near the top. Each company keeps its own projects, sections, and tasks, so pick the one you plan to update.",
  "Open the Project dropdown to jump into work inside that company. Use the + button to create new projects, or the pencil/bin icons beside a project to rename or delete it.",
  "Keep sections organised. Click Add section or use the menu on each column to rename or remove it, and drag section headers to change their order.",
  "Add tasks with the Quick Add form or from the board/list view. Set the project, section, due date, assignee, and priority, then drag cards between sections as work moves forward.",
  "Need to move work between companies? Edit the task or project and pick the new project - its company automatically follows, and tasks keep their history.",
  "Keep everyone aligned: update this Userguide whenever the process changes (Userguide -> Edit guide) and export tasks regularly if you need an offline backup.",
];
const LEGACY_USERGUIDES = [LEGACY_USERGUIDE_V1, LEGACY_USERGUIDE_V2];

const state = {
  tasks: [],
  projects: [],
  sections: [],
  companies: [],
  members: [],
  departments: [],
  settings: null,
  activeView: { type: "view", value: "workspace" },
  activeCompanyId: DEFAULT_COMPANY.id,
  companyRecents: {},
  userguide: [...DEFAULT_USERGUIDE],
  viewMode: "list",
  searchTerm: "",
  activeTaskTab: "active",
  activeAllTab: "created",
  metricsFilter: "all",
  editingTaskId: null,
  editingProjectId: null,
  editingCompanyId: null,
  editingMeetingId: null,
  editingEmailId: null,
  meetingActionDraft: [],
  emailThreadDraft: [],
  meetingCompletedDraft: false,
  emailCompletedDraft: false,
  textareaHeights: {},
  dragTaskId: null,
  dragSectionId: null,
  sectionDropTarget: null,
  isQuickAddOpen: false,
  isUserguideOpen: false,
  isEditingUserguide: false,
  openDropdown: null,
  dialogAttachmentDraft: [],
  filters: {
    member: "",
    department: "",
    priority: "",
    due: "",
  },
  projectSectionLimits: {},
  recentActivityLimit: 10,
  importJob: {
    file: null,
    status: "idle",
    error: "",
    model: OPENROUTER_MODEL,
    provider: WHATSAPP_DEFAULT_PROVIDER,
    stats: null,
    companyId: DEFAULT_COMPANY.id,
    projectId: null,
  },
  imports: {
    whatsapp: {},
  },
};

const elements = {
  viewList: document.getElementById("viewList"),
  viewCounts: {
    workspace: document.querySelector('[data-count="workspace"]'),
    today: document.querySelector('[data-count="today"]'),
    upcoming: document.querySelector('[data-count="upcoming"]'),
  },
  searchInput: document.getElementById("searchInput"),
  searchInputMobile: document.getElementById("searchInput-mobile"),
  searchResults: document.getElementById("searchResults"),
  profileAvatar: document.getElementById("profileAvatar"),
  profileNameDisplay: document.getElementById("profileNameDisplay"),
  metricFilter: document.getElementById("metricFilter"),
  addProject: document.getElementById("addProject"),
  quickAddForm: document.getElementById("quickAddForm"),
  quickAddProject: document.querySelector('#quickAddForm select[name="project"]'),
  quickAddPriority: document.querySelector('#quickAddForm select[name="priority"]'),
  quickAddSection: document.querySelector('#quickAddForm select[name="section"]'),
  quickAddDepartment: document.querySelector('#quickAddForm select[name="department"]'),
  quickAddAssignee: document.querySelector('#quickAddForm select[name="assignee"]'),
  quickAddAttachments: document.querySelector('#quickAddForm input[name="attachments"]'),
  quickAddAttachmentList: document.querySelector('#quickAddForm [data-attachment-list]'),
  quickAddLinksList: document.querySelector("[data-quick-links]"),
  quickAddCancel: document.querySelector('#quickAddForm [data-action="cancel"]'),
  quickAddError: document.querySelector(".quick-add-error"),
  toggleQuickAdd: document.getElementById("toggleQuickAdd"),
  sectionActions: document.getElementById("sectionActions"),
  addSection: document.getElementById("addSection"),
  boardView: document.getElementById("boardView"),
  boardColumns: document.getElementById("boardColumns"),
  boardEmptyState: document.getElementById("boardEmptyState"),
  listView: document.getElementById("listView"),
  taskTabList: document.getElementById("taskTabList"),
  taskTabPanels: document.getElementById("taskTabPanels"),
  viewTitle: document.getElementById("viewTitle"),
  viewSubtitle: document.getElementById("viewSubtitle"),
  viewToggleButtons: [...document.querySelectorAll('[data-view-mode]')],
  manageMembers: document.getElementById("manageMembers"),
  manageDepartments: document.getElementById("manageDepartments"),
  membersDialog: document.getElementById("membersDialog"),
  membersForm: document.getElementById("membersForm"),
  memberList: document.getElementById("memberList"),
  departmentsDialog: document.getElementById("departmentsDialog"),
  departmentsForm: document.getElementById("departmentsForm"),
  departmentList: document.getElementById("departmentList"),
  exportTasks: document.getElementById("exportTasks"),
  taskDialog: document.getElementById("taskDialog"),
  dialogForm: document.getElementById("dialogForm"),
  taskDialogStatus: document.querySelector("[data-task-status]"),
  taskDialogToggle: document.querySelector('[data-action="toggle-task-completion"]'),
  dialogAttachmentsInput: document.querySelector('#dialogForm input[name="attachments"]'),
  dialogAttachmentList: document.querySelector('#dialogForm [data-dialog-attachment-list]'),
  dialogLinksList: document.querySelector("[data-dialog-links]"),
  taskTemplate: document.getElementById("taskItemTemplate"),
  activeTasksMetric: document.getElementById("active-tasks"),
  activityFeed: document.getElementById("activity-feed"),
  filterBar: document.getElementById("taskFilters"),
  filterMember: document.getElementById("filterMember"),
  filterDepartment: document.getElementById("filterDepartment"),
  filterPriority: document.getElementById("filterPriority"),
  filterDue: document.getElementById("filterDue"),
  filterClear: document.querySelector('[data-action="clear-filters"]'),
  userguidePanel: document.getElementById("userguidePanel"),
  companyTabs: document.getElementById("companyTabs"),
  companyTabsEmpty: document.getElementById("companyTabsEmpty"),
  projectDropdownToggle: document.getElementById("projectDropdownToggle"),
  projectDropdownMenu: document.getElementById("projectDropdownMenu"),
  projectDropdownList: document.getElementById("projectDropdownList"),
  projectDropdownLabel: document.getElementById("projectDropdownLabel"),
  meetingDialog: document.getElementById("meetingDialog"),
  meetingForm: document.getElementById("meetingForm"),
  meetingProjectLabel: document.querySelector("[data-meeting-project]"),
  meetingDialogStatus: document.querySelector("[data-meeting-status]"),
  meetingDialogToggle: document.querySelector('[data-action="meeting-toggle-completion"]'),
  meetingDepartment: document.querySelector('#meetingForm select[name="department"]'),
  meetingPriority: document.querySelector('#meetingForm select[name="priority"]'),
  meetingAssignee: document.querySelector('#meetingForm select[name="assignee"]'),
  meetingLinksList: document.querySelector("[data-meeting-links]"),
  meetingError: document.querySelector("[data-meeting-error]"),
  meetingActionList: document.querySelector("[data-meeting-action-list]"),
  meetingActionInput: document.querySelector("[data-meeting-action-input]"),
  emailDialog: document.getElementById("emailDialog"),
  emailForm: document.getElementById("emailForm"),
  emailProjectLabel: document.querySelector("[data-email-project]"),
  emailDialogStatus: document.querySelector("[data-email-status]"),
  emailDialogToggle: document.querySelector('[data-action="email-toggle-completion"]'),
  emailDepartment: document.querySelector('#emailForm select[name="department"]'),
  emailPriority: document.querySelector('#emailForm select[name="priority"]'),
  emailStatus: document.querySelector('#emailForm select[name="status"]'),
  emailAssignee: document.querySelector('#emailForm select[name="assignee"]'),
  emailLinksList: document.querySelector("[data-email-links]"),
  emailThreadList: document.querySelector("[data-email-thread-list]"),
  emailThreadInput: document.querySelector("[data-email-thread-input]"),
  emailError: document.querySelector("[data-email-error]"),
  companyDialog: document.getElementById("companyDialog"),
  companyForm: document.getElementById("companyForm"),
  companyNameInput: document.getElementById("companyName"),
  companyError: document.querySelector("[data-company-error]"),
  projectDialog: document.getElementById("projectDialog"),
  projectForm: document.getElementById("projectForm"),
  projectNameInput: document.querySelector('#projectForm input[name="projectName"]'),
  projectCompanySelect: document.querySelector('#projectForm select[name="projectCompany"]'),
  projectError: document.querySelector("[data-project-error]"),
  userguideList: document.getElementById("userguideList"),
  userguideForm: document.getElementById("userguideForm"),
  userguideEditor: document.getElementById("userguideEditor"),
  userguideEditToggle: document.querySelector('[data-action="edit-userguide"]'),
  userguideCancelEdit: document.querySelector('[data-action="cancel-userguide"]'),
  importWhatsapp: document.getElementById("importWhatsapp"),
  whatsappDialog: document.getElementById("whatsappDialog"),
  whatsappForm: document.getElementById("whatsappForm"),
  whatsappFile: document.getElementById("whatsappFile"),
  whatsappPreview: document.getElementById("whatsappPreview"),
  whatsappFileLabel: document.querySelector("[data-whatsapp-file]"),
  whatsappRangeLabel: document.querySelector("[data-whatsapp-range]"),
  whatsappSummary: document.querySelector("[data-whatsapp-summary]"),
  whatsappError: document.getElementById("whatsappError"),
  whatsappModel: document.getElementById("whatsappModel"),
  whatsappProject: document.getElementById("whatsappProject"),
  whatsappCompany: document.querySelector("[data-whatsapp-company]"),
};


const loadJSON = (key, fallback) => {
  try {
    const raw = window.localStorage.getItem(key);
    if (!raw) return fallback;
    const parsed = JSON.parse(raw);
    return parsed ?? fallback;
  } catch {
    return fallback;
  }
};

const saveJSON = (key, value) => {
  try {
    window.localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Ignore quota failures (private mode, quota full, etc.).
  }
};

const saveTasks = () => {
  saveJSON(STORAGE_KEYS.tasks, state.tasks);
  persistWorkspace();
};
const saveProjects = () => {
  saveJSON(STORAGE_KEYS.projects, state.projects);
  persistWorkspace();
};
const saveSections = () => {
  saveJSON(STORAGE_KEYS.sections, state.sections);
  persistWorkspace();
};
const saveCompanies = () => {
  saveJSON(STORAGE_KEYS.companies, state.companies);
  persistWorkspace();
};
const saveMembers = () => {
  saveJSON(STORAGE_KEYS.members, state.members);
  persistWorkspace();
};
const saveDepartments = () => {
  saveJSON(STORAGE_KEYS.departments, state.departments);
  persistWorkspace();
};
const saveUserguide = () => {
  saveJSON(STORAGE_KEYS.userguide, state.userguide);
  persistWorkspace();
};
const saveImports = () => {
  saveJSON(STORAGE_KEYS.imports, state.imports);
  persistWorkspace();
};
const savePreferences = () =>
  saveJSON(STORAGE_KEYS.preferences, {
    activeView: state.activeView,
    viewMode: state.viewMode,
    activeCompanyId: state.activeCompanyId,
    companyRecents: state.companyRecents,
    activeTaskTab: state.activeTaskTab,
    activeAllTab: state.activeAllTab,
    metricsFilter: state.metricsFilter,
    textareaHeights: state.textareaHeights,
  });

const applyStoredPreferences = () => {
  const prefs = loadJSON(STORAGE_KEYS.preferences, {});
  if (prefs.activeView?.type && prefs.activeView?.value) {
    state.activeView = prefs.activeView;
  }
  if (prefs.viewMode === "board" || prefs.viewMode === "list") {
    state.viewMode = prefs.viewMode;
  }
  if (typeof prefs.activeCompanyId === "string") {
    state.activeCompanyId = prefs.activeCompanyId;
  }
  if (prefs.companyRecents && typeof prefs.companyRecents === "object") {
    state.companyRecents = { ...prefs.companyRecents };
  }
  if (typeof prefs.activeTaskTab === "string") {
    state.activeTaskTab = prefs.activeTaskTab;
  }
  if (typeof prefs.activeAllTab === "string") {
    state.activeAllTab = prefs.activeAllTab;
  }
  if (typeof prefs.metricsFilter === "string") {
    state.metricsFilter = prefs.metricsFilter;
  }
  if (prefs.textareaHeights && typeof prefs.textareaHeights === "object") {
    state.textareaHeights = { ...prefs.textareaHeights };
  }
  if (state.textareaHeights) {
    AUTOSIZE_RESET_KEYS.forEach((key) => {
      if (key in state.textareaHeights) {
        delete state.textareaHeights[key];
      }
    });
  }

  if (state.viewMode === "board" && state.activeView.type !== "project") {
    state.viewMode = "list";
  }
  ensureCompanyPreferences();
};

const generateId = (prefix) => {
  const fallback = `${prefix}-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
  return window.crypto?.randomUUID?.() ?? fallback;
};

const pickProjectColor = (index) => PROJECT_COLORS[index % PROJECT_COLORS.length];

const getCompanyById = (companyId) => state.companies.find((company) => company.id === companyId);
const getProjectById = (projectId) => state.projects.find((project) => project.id === projectId);
const getSectionById = (sectionId) => state.sections.find((section) => section.id === sectionId);
const getSectionsForProject = (projectId) =>
  state.sections
    .filter((section) => section.projectId === projectId)
    .sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
const getProjectsForCompany = (companyId) =>
  state.projects.filter((project) => project.companyId === companyId && !project.isDefault);
const tasksForCompany = (companyId = state.activeCompanyId) =>
  state.tasks.filter((task) => {
    if (task.deletedAt) return false;
    if (!companyId || companyId === ALL_COMPANY_ID) return true;
    if (!task.companyId) return companyId === DEFAULT_COMPANY.id;
    return task.companyId === companyId;
  });

const rememberProjectSelection = (projectId) => {
  const project = getProjectById(projectId);
  if (!project) return;
  state.activeCompanyId = project.companyId;
  state.companyRecents[project.companyId] = project.id;
};

const getPreferredProjectId = () => {
  if (state.activeView.type === "project" && getProjectById(state.activeView.value)) {
    return state.activeView.value;
  }
  const remembered = state.companyRecents[state.activeCompanyId];
  if (remembered && getProjectById(remembered)) {
    return remembered;
  }
  const fallback = getProjectsForCompany(state.activeCompanyId)[0];
  if (fallback) {
    return fallback.id;
  }
  return DEFAULT_PROJECT.id;
};

const getMemberById = (memberId) => state.members.find((member) => member.id === memberId);
const getDepartmentById = (departmentId) =>
  state.departments.find((department) => department.id === departmentId);

const ensureDefaultCompany = () => {
  if (!state.companies.length) {
    state.companies = [{ ...DEFAULT_COMPANY }];
    return;
  }
  const hasDefault = state.companies.some((company) => company.id === DEFAULT_COMPANY.id);
  if (!hasDefault) {
    state.companies.unshift({ ...DEFAULT_COMPANY });
  }
};

const ensureDefaultProject = () => {
  const index = state.projects.findIndex((project) => project.id === DEFAULT_PROJECT.id);
  if (index === -1) {
    state.projects.unshift({ ...DEFAULT_PROJECT });
    return;
  }
  if (!state.projects[index].companyId) {
    state.projects[index] = { ...state.projects[index], companyId: DEFAULT_COMPANY.id };
  }
};

const ensureProjectsHaveCompany = () => {
  let mutated = false;
  state.projects = state.projects.map((project) => {
    if (project.companyId && getCompanyById(project.companyId)) {
      return project;
    }
    mutated = true;
    return { ...project, companyId: DEFAULT_COMPANY.id };
  });
  if (mutated) {
    const seen = new Map();
    state.projects.forEach((project) => {
      seen.set(project.id, project.companyId);
    });
    state.tasks = state.tasks.map((task) => {
      if (!task.projectId) return task;
      const companyId = seen.get(task.projectId);
      if (!companyId) return task;
      return { ...task, companyId };
    });
  }
};

const ensureCompanyPreferences = () => {
  if (!getCompanyById(state.activeCompanyId)) {
    state.activeCompanyId = state.companies[0]?.id ?? DEFAULT_COMPANY.id;
  }
  const validProjects = new Map(state.projects.map((project) => [project.id, project.companyId]));
  const nextRecents = {};
  Object.entries(state.companyRecents || {}).forEach(([companyId, projectId]) => {
    if (validProjects.get(projectId) === companyId) {
      nextRecents[companyId] = projectId;
    }
  });
  state.companyRecents = nextRecents;
  if (!state.companyRecents[state.activeCompanyId]) {
    const fallbackProject = state.projects.find(
      (project) => project.companyId === state.activeCompanyId && !project.isDefault,
    );
    if (fallbackProject) {
      state.companyRecents[state.activeCompanyId] = fallbackProject.id;
    }
  }
};

const normaliseUserguide = (entries) => {
  if (!Array.isArray(entries)) return [...DEFAULT_USERGUIDE];
  const cleaned = entries
    .map((entry) => (typeof entry === "string" ? entry.trim() : ""))
    .filter(Boolean);
  return cleaned.length ? cleaned : [...DEFAULT_USERGUIDE];
};

const userguideEquals = (a, b) => {
  if (!Array.isArray(a) || !Array.isArray(b)) return false;
  if (a.length !== b.length) return false;
  return a.every((entry, index) => entry.trim() === (b[index] ?? "").trim());
};

const isLegacyUserguide = (entries) => LEGACY_USERGUIDES.some((legacy) => userguideEquals(entries, legacy));

const upgradeUserguideIfLegacy = () => {
  if (!isLegacyUserguide(state.userguide)) return;
  state.userguide = [...DEFAULT_USERGUIDE];
  saveJSON(STORAGE_KEYS.userguide, state.userguide);
  if (remoteLoaded) {
    saveUserguide();
  } else {
    pendingUserguideUpgrade = true;
  }
};

const ensureUserguideHasLatestEntries = () => {
  if (!Array.isArray(state.userguide)) return;
  let mutated = false;
  USERGUIDE_LATEST_ENTRIES.forEach((entry) => {
    if (!state.userguide.some((line) => line.trim() === entry.trim())) {
      state.userguide.push(entry);
      mutated = true;
    }
  });
  if (mutated) {
    saveJSON(STORAGE_KEYS.userguide, state.userguide);
    if (remoteLoaded) {
      saveUserguide();
    } else {
      pendingUserguideUpgrade = true;
    }
  }
};

const setActiveCompany = (companyId) => {
  if (companyId === ALL_COMPANY_ID) {
    if (state.activeCompanyId === ALL_COMPANY_ID) return;
    state.activeCompanyId = ALL_COMPANY_ID;
    state.activeView = { type: "view", value: "workspace" };
    state.activeTaskTab = "active";
    state.activeAllTab = "created";
    if (state.viewMode === "board") {
      state.viewMode = "list";
    }
    savePreferences();
    render();
    return;
  }

  const company = getCompanyById(companyId);
  if (!company || state.activeCompanyId === companyId) return;
  state.activeCompanyId = companyId;
  state.activeTaskTab = "active";
  const remembered = state.companyRecents[companyId];
  const rememberedProject = remembered ? getProjectById(remembered) : null;
  if (rememberedProject) {
    setActiveView("project", rememberedProject.id);
    return;
  }
  const fallback = getProjectsForCompany(companyId)[0];
  if (fallback) {
    setActiveView("project", fallback.id);
    return;
  }
  if (state.activeView.type === "project") {
    setActiveView("view", "workspace");
  } else {
    savePreferences();
    render();
  }
};

const ensureDefaultDepartment = () => {
  const hasDefault = state.departments.some((department) => department.id === DEFAULT_DEPARTMENT.id);
  if (!hasDefault) {
    state.departments.unshift({ ...DEFAULT_DEPARTMENT });
  }
};

const ensureSectionForProject = (projectId) => {
  let section = getSectionsForProject(projectId)[0];
  if (!section) {
    section = {
      id: generateId("section"),
      name: DEFAULT_SECTION_NAME,
      projectId,
      order: 0,
      createdAt: new Date().toISOString(),
    };
    state.sections.push(section);
    saveSections();
  }
  return section;
};

const ensureAllProjectsHaveSections = () => {
  let addedSection = false;
  state.projects.forEach((project) => {
    const hasSection = state.sections.some((section) => section.projectId === project.id);
    if (!hasSection) {
      state.sections.push({
        id: generateId("section"),
        name: DEFAULT_SECTION_NAME,
        projectId: project.id,
        order: 0,
        createdAt: new Date().toISOString(),
      });
      addedSection = true;
    }
  });
  if (addedSection) {
    saveSections();
  }
};

const defaultSettings = () => ({
  profile: {
    name: "Sarah Chen",
    photo: "https://images.unsplash.com/photo-1472099645785-5658abf4ff4e?w=96&h=96&fit=crop&crop=face",
  },
  theme: {
    accent: "#2563eb",
    priorities: {
      critical: "#dc2626",
      veryHigh: "#ea580c",
      high: "#f97316",
      medium: "#0ea5e9",
      low: "#10b981",
      optional: "#6366f1",
    },
  },
  display: {
    fontScale: 1,
  },
});

const clampFontScale = (value) => {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 1;
  if (numeric < 0.9) return 0.9;
  if (numeric > 1.15) return 1.15;
  return numeric;
};

const normaliseSettings = (settings = {}) => {
  const defaults = defaultSettings();
  const merged = {
    profile: {
      name: settings.profile?.name?.trim() || defaults.profile.name,
      photo: settings.profile?.photo || defaults.profile.photo,
    },
    theme: {
      accent: settings.theme?.accent || defaults.theme.accent,
      priorities: {
        critical: settings.theme?.priorities?.critical || defaults.theme.priorities.critical,
        veryHigh: settings.theme?.priorities?.veryHigh || defaults.theme.priorities.veryHigh,
        high: settings.theme?.priorities?.high || defaults.theme.priorities.high,
        medium: settings.theme?.priorities?.medium || defaults.theme.priorities.medium,
        low: settings.theme?.priorities?.low || defaults.theme.priorities.low,
        optional: settings.theme?.priorities?.optional || defaults.theme.priorities.optional,
      },
    },
    display: {
      fontScale: clampFontScale(settings.display?.fontScale || defaults.display.fontScale),
    },
  };
  return merged;
};

const WORKSPACE_ID = import.meta.env.VITE_FIREBASE_WORKSPACE_ID ?? 'default';
const workspaceRef = doc(db, 'workspaces', WORKSPACE_ID);

let workspaceUnsubscribe = null;
let remoteLoaded = false;
let suppressSnapshot = false;
let pendingUserguideUpgrade = false;

const getStateSnapshot = () => ({
  tasks: state.tasks,
  projects: state.projects,
  sections: state.sections,
  companies: state.companies,
  members: state.members,
  departments: state.departments,
  settings: state.settings,
  userguide: state.userguide,
  imports: state.imports,
});

const persistWorkspace = async () => {
  if (!remoteLoaded) return;
  try {
    suppressSnapshot = true;
    await setDoc(
      workspaceRef,
      {
        ...getStateSnapshot(),
        updatedAt: new Date().toISOString(),
      },
      { merge: true },
    );
  } catch (error) {
    console.error('Failed to persist workspace to Firestore', error);
  } finally {
    suppressSnapshot = false;
  }
};

const applyRemoteState = (data) => {
  const incoming = data || {};
  state.tasks = Array.isArray(incoming.tasks) ? incoming.tasks : [];
  state.projects = Array.isArray(incoming.projects) ? incoming.projects : [];
  state.sections = Array.isArray(incoming.sections) ? incoming.sections : [];
  state.companies = Array.isArray(incoming.companies) ? incoming.companies : [];
  state.members = Array.isArray(incoming.members) ? incoming.members : [];
  state.departments = Array.isArray(incoming.departments) ? incoming.departments : [];
  state.settings = normaliseSettings(incoming.settings ?? defaultSettings());
  state.userguide = normaliseUserguide(incoming.userguide);
  const incomingImports = incoming.imports;
  state.imports = {
    whatsapp: { ...(incomingImports?.whatsapp ?? {}) },
  };
  upgradeUserguideIfLegacy();
  ensureUserguideHasLatestEntries();

  ensureDefaultCompany();
  ensureDefaultProject();
  ensureProjectsHaveCompany();
  ensureDefaultDepartment();
  ensureAllProjectsHaveSections();
  ensureCompanyPreferences();
  state.tasks.forEach(ensureTaskDefaults);

  saveJSON(STORAGE_KEYS.tasks, state.tasks);
  saveJSON(STORAGE_KEYS.projects, state.projects);
  saveJSON(STORAGE_KEYS.sections, state.sections);
  saveJSON(STORAGE_KEYS.companies, state.companies);
  saveJSON(STORAGE_KEYS.members, state.members);
  saveJSON(STORAGE_KEYS.departments, state.departments);
  saveJSON(STORAGE_KEYS.settings, state.settings);
  saveJSON(STORAGE_KEYS.userguide, state.userguide);
  saveJSON(STORAGE_KEYS.imports, state.imports);

  applySettings();
};

const startWorkspaceSync = async () => {
  try {
    const snapshot = await getDoc(workspaceRef);
    if (snapshot.exists()) {
      applyRemoteState(snapshot.data());
    } else {
      await setDoc(workspaceRef, {
        ...getStateSnapshot(),
        updatedAt: new Date().toISOString(),
      });
    }
    remoteLoaded = true;
    if (pendingUserguideUpgrade) {
      pendingUserguideUpgrade = false;
      saveUserguide();
    }
    applyStoredPreferences();
    render();

    workspaceUnsubscribe = onSnapshot(
      workspaceRef,
      (docSnapshot) => {
        if (suppressSnapshot) return;
        if (!docSnapshot.exists()) return;
        applyRemoteState(docSnapshot.data());
        render();
      },
      (error) => {
        console.error('Firestore realtime listener error', error);
      },
    );
  } catch (error) {
    console.error('Failed to load workspace from Firestore', error);
    throw error;
  }
};

const hexToRgba = (hex, alpha) => {
  if (!hex) return `rgba(37, 99, 235, ${alpha})`;
  const normalized = (hex || '').replace('#', '');
  const extended = normalized.length === 3
    ? normalized.split('').map((char) => char + char).join('')
    : normalized;
  const bigint = Number.parseInt(extended || '000000', 16);
  const r = (bigint >> 16) & 255;
  const g = (bigint >> 8) & 255;
  const b = bigint & 255;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
};

const isSameDay = (isoString, reference = new Date()) => {
  if (!isoString) return false;
  const date = new Date(isoString);
  if (Number.isNaN(date.getTime())) return false;
  const ref = new Date(reference);
  return (
    date.getFullYear() === ref.getFullYear() &&
    date.getMonth() === ref.getMonth() &&
    date.getDate() === ref.getDate()
  );
};

const getDefaultSectionId = (projectId) => ensureSectionForProject(projectId).id;

const normaliseTitle = (value) => value.trim();

const readFilesAsData = async (fileList) => {
  const files = Array.from(fileList || []);
  if (!files.length) return [];
  const readers = files.map((file) => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve({
      id: generateId("file"),
      name: file.name,
      type: file.type,
      size: file.size,
      data: reader.result,
    });
    reader.onerror = () => reject(new Error(`Failed to read attachment: ${file.name}`));
    reader.readAsDataURL(file);
  }));
  return Promise.all(readers);
};

const createAttachmentChip = (attachment) => {
  const link = document.createElement("a");
  link.className = "attachment-chip";
  link.href = attachment.data;
  link.download = attachment.name;
  link.target = "_blank";
  link.rel = "noopener";
  link.textContent = attachment.name;
  return link;
};


const ensureTaskDefaults = (task) => {
  if (!getProjectById(task.projectId)) {
    task.projectId = "inbox";
  }
  const project = getProjectById(task.projectId);
  task.companyId = project?.companyId ?? DEFAULT_COMPANY.id;
  const validSection = task.sectionId && getSectionById(task.sectionId);
  if (!validSection) {
    task.sectionId = getDefaultSectionId(task.projectId);
  }
  if (!state.departments.some((department) => department.id === task.departmentId)) {
    task.departmentId = "";
  }
  if (!state.members.some((member) => member.id === task.assigneeId)) {
    task.assigneeId = "";
  }
  if (typeof task.priority !== "string") {
    task.priority = "medium";
  }
  if (!task.createdAt) {
    task.createdAt = new Date().toISOString();
  }
  if (!task.updatedAt) {
    task.updatedAt = task.createdAt;
  }
  if (typeof task.deletedAt !== "string") {
    task.deletedAt = null;
  }
  if (!Array.isArray(task.attachments)) {
    task.attachments = [];
  }
  task.kind = normaliseTaskKind(task.kind);
  task.source = normaliseTaskSource(task.source);
  if (!task.metadata || typeof task.metadata !== "object") {
    task.metadata = {};
  }
  task.links = normaliseTaskLinks(task.links);
  const rawActionItems = Array.isArray(task.metadata.actionItems)
    ? task.metadata.actionItems
    : Array.isArray(task.actionItems)
      ? task.actionItems
      : [];
  const actionItems = rawActionItems
    .map((item) => normaliseActionItem(item))
    .filter(Boolean);
  task.metadata.actionItems = actionItems;
  task.actionItems = actionItems;
  const rawThreadMessages = Array.isArray(task.metadata.threadMessages)
    ? task.metadata.threadMessages
    : [];
  const threadMessages = rawThreadMessages
    .map((entry) => normaliseThreadMessage(entry))
    .filter(Boolean);
  task.metadata.threadMessages = threadMessages;
};
const describeDueDate = (dueDate) => {
  if (!dueDate) return { label: "", className: "" };
  const date = new Date(dueDate);
  if (Number.isNaN(date.getTime())) return { label: "", className: "" };

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const due = new Date(date);
  due.setHours(0, 0, 0, 0);

  const diffMs = due.getTime() - today.getTime();
  const diffDays = Math.round(diffMs / (1000 * 60 * 60 * 24));

  if (diffDays === 0) return { label: "Today", className: "meta-chip today" };
  if (diffDays === 1) return { label: "Tomorrow", className: "meta-chip today" };
  if (diffDays < 0) {
    const formatted = new Intl.DateTimeFormat(undefined, { month: "short", day: "numeric" }).format(date);
    return { label: `Due ${formatted}`, className: "meta-chip overdue" };
  }

  const formatted = new Intl.DateTimeFormat(undefined, {
    weekday: diffDays < 7 ? "short" : undefined,
    month: "short",
    day: "numeric",
  }).format(date);
  return { label: formatted, className: "meta-chip" };
};

const getTaskDescriptionText = (task) => {
  if (typeof task.description === "string" && task.description.trim()) {
    return task.description.trim();
  }
  if (typeof task.metadata?.notes === "string" && task.metadata.notes.trim()) {
    return task.metadata.notes.trim();
  }
  if (task.kind === "meeting" && typeof task.metadata?.attendees === "string") {
    const attendees = task.metadata.attendees.trim();
    if (attendees) {
      return `Attendees: ${attendees}`;
    }
  }
  return "";
};

const buildPreviewFromText = (text, lines = 3) => {
  if (!text) return "";
  const segments = text
    .split(/\r?\n/)
    .map((segment) => segment.trim())
    .filter(Boolean);
  if (!segments.length) return "";
  return segments.slice(0, lines).join("\n");
};

const describeTaskPreview = (task) => {
  const preview = buildPreviewFromText(getTaskDescriptionText(task));
  return preview || "No notes yet.";
};

const describeTaskAssignee = (task) => {
  const member = getMemberById(task.assigneeId);
  return member ? `Assignee: ${member.name}` : "Assignee: Unassigned";
};

const describeTaskActionItems = (task) => {
  const items = Array.isArray(task.metadata?.actionItems) ? task.metadata.actionItems : [];
  if (!items.length) return "Action items: none";
  const remaining = items.filter((item) => !item?.completed).length;
  if (remaining === 0) {
    return `Action items: ${items.length} complete`;
  }
  return `Action items: ${remaining}/${items.length} open`;
};

const describeTaskDueLabel = (task) => {
  const { label } = describeDueDate(task.dueDate);
  return label ? `Due ${label}` : "No due date";
};

const describeTaskCompletionStatus = (task) => {
  if (!task.completed) return "In progress";
  if (!task.completedAt) return "Completed";
  const completedDate = new Date(task.completedAt);
  if (Number.isNaN(completedDate.getTime())) return "Completed";
  const formatted = new Intl.DateTimeFormat(undefined, {
    month: "short",
    day: "numeric",
    year: completedDate.getFullYear() !== new Date().getFullYear() ? "numeric" : undefined,
  }).format(completedDate);
  return `Completed ${formatted}`;
};

const updateTaskDialogCompletionState = (task) => {
  if (!elements.taskDialogToggle) return;
  const completed = Boolean(task.completed);
  elements.taskDialogToggle.textContent = completed ? "Restore task" : "Mark completed";
  elements.taskDialogToggle.dataset.completedState = completed ? "true" : "false";
  if (elements.taskDialogStatus) {
    elements.taskDialogStatus.textContent = describeTaskCompletionStatus(task);
  }
  if (elements.dialogForm?.elements?.completed) {
    elements.dialogForm.elements.completed.checked = completed;
  }
};

const updateMeetingDialogCompletionState = (completed, completedAt = null) => {
  if (elements.meetingDialogToggle) {
    elements.meetingDialogToggle.dataset.completedState = completed ? "true" : "false";
    elements.meetingDialogToggle.textContent = completed ? "Restore meeting" : "Mark completed";
  }
  if (elements.meetingDialogStatus) {
    elements.meetingDialogStatus.textContent = describeTaskCompletionStatus({
      completed,
      completedAt,
    });
  }
};

const updateEmailDialogCompletionState = (completed, completedAt = null) => {
  if (elements.emailDialogToggle) {
    elements.emailDialogToggle.dataset.completedState = completed ? "true" : "false";
    elements.emailDialogToggle.textContent = completed ? "Restore email" : "Mark completed";
  }
  if (elements.emailDialogStatus) {
    elements.emailDialogStatus.textContent = describeTaskCompletionStatus({
      completed,
      completedAt,
    });
  }
};

const compareTasks = (a, b) => {
  const dueA = a.dueDate ? new Date(a.dueDate).getTime() : Number.POSITIVE_INFINITY;
  const dueB = b.dueDate ? new Date(b.dueDate).getTime() : Number.POSITIVE_INFINITY;
  if (dueA !== dueB) return dueA - dueB;
  const weightA = PRIORITY_WEIGHT[a.priority] ?? 1;
  const weightB = PRIORITY_WEIGHT[b.priority] ?? 1;
  if (weightA !== weightB) return weightA - weightB;
  return new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime();
};

const matchesActiveView = (task) => {
  const { type, value } = state.activeView;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const matchesCompany =
    !state.activeCompanyId ||
    state.activeCompanyId === ALL_COMPANY_ID ||
    task.companyId === state.activeCompanyId ||
    (!task.companyId && state.activeCompanyId === DEFAULT_COMPANY.id);
  if (!matchesCompany) return false;

  if (type === "view") {
    if (value === "workspace") return task.projectId === "inbox";
    if (value === "today") {
      if (!task.dueDate) return false;
      const due = new Date(task.dueDate);
      due.setHours(0, 0, 0, 0);
      return due.getTime() === today.getTime();
    }
    if (value === "upcoming") {
      if (!task.dueDate) return false;
      const due = new Date(task.dueDate);
      due.setHours(0, 0, 0, 0);
      return due.getTime() > today.getTime();
    }
  }
  if (type === "project") return task.projectId === value;
  return true;
};

const matchesSearch = (task) => {
  const needle = state.searchTerm.trim().toLowerCase();
  if (!needle) return true;

  const section = getSectionById(task.sectionId);
  const project = getProjectById(task.projectId);
  const member = getMemberById(task.assigneeId);
  const department = getDepartmentById(task.departmentId);
  const attachments = Array.isArray(task.attachments) ? task.attachments : [];
  const links = Array.isArray(task.links) ? task.links : [];

  const metadataValues = [];
  if (task.metadata && typeof task.metadata === "object") {
    Object.values(task.metadata).forEach((value) => {
      if (typeof value === "string") {
        metadataValues.push(value.toLowerCase());
      } else if (Array.isArray(value)) {
        value.forEach((entry) => {
          if (typeof entry === "string") {
            metadataValues.push(entry.toLowerCase());
          } else if (entry && typeof entry === "object") {
            Object.values(entry).forEach((nested) => {
              if (typeof nested === "string") {
                metadataValues.push(nested.toLowerCase());
              }
            });
          }
        });
      } else if (value && typeof value === "object") {
        Object.values(value).forEach((nested) => {
          if (typeof nested === "string") {
            metadataValues.push(nested.toLowerCase());
          }
        });
      }
    });
  }

  return (
    task.title.toLowerCase().includes(needle) ||
    (task.description ?? "").toLowerCase().includes(needle) ||
    (task.kind ?? "").toString().toLowerCase().includes(needle) ||
    (task.source ?? "").toString().toLowerCase().includes(needle) ||
    (section?.name ?? "").toLowerCase().includes(needle) ||
    (member?.name ?? "").toLowerCase().includes(needle) ||
    (department?.name ?? "").toLowerCase().includes(needle) ||
    (project?.name ?? "").toLowerCase().includes(needle) ||
    attachments.some((attachment) => (attachment.name ?? "").toLowerCase().includes(needle)) ||
    links.some((link) =>
      (link.title ?? "").toLowerCase().includes(needle) ||
      (link.url ?? "").toLowerCase().includes(needle)
    ) ||
    metadataValues.some((entry) => entry.includes(needle))
  );
};

const isFilterActive = () => Object.values(state.filters || {}).some(Boolean);

const matchesFilters = (task) => {
  if (!state.filters) return true;
  const { member, department, priority, due } = state.filters;
  if (member) {
    if (member === FILTER_MEMBER_UNASSIGNED) {
      if (task.assigneeId) return false;
    } else if (task.assigneeId !== member) {
      return false;
    }
  }
  if (department && task.departmentId !== department) {
    return false;
  }
  if (priority) {
    const taskPriority = typeof task.priority === "string" ? task.priority : "medium";
    if (taskPriority !== priority) {
      return false;
    }
  }
  if (due) {
    const dueDate = task.dueDate ? new Date(task.dueDate) : null;
    const hasValidDate = dueDate && !Number.isNaN(dueDate.getTime());
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    if (due === "none") {
      if (hasValidDate) return false;
    } else if (!hasValidDate) {
      return false;
    } else {
      const dueDay = new Date(dueDate);
      dueDay.setHours(0, 0, 0, 0);
      if (due === "overdue") {
        if (!(dueDay < today && !task.completed)) return false;
      } else if (due === "today") {
        if (dueDay.getTime() !== today.getTime()) return false;
      } else if (due === "week") {
        const weekEnd = new Date(today);
        weekEnd.setDate(today.getDate() + 7);
        if (dueDay < today || dueDay > weekEnd) return false;
      }
    }
  }
  return true;
};

const setSearchTerm = (value) => {
  const trimmed = value.trim();
  state.searchTerm = trimmed;
  if (elements.searchInput && elements.searchInput.value !== trimmed) {
    elements.searchInput.value = trimmed;
  }
  if (elements.searchInputMobile && elements.searchInputMobile.value !== trimmed) {
    elements.searchInputMobile.value = trimmed;
  }
  if (!trimmed) {
    hideSearchResults();
  } else {
    renderSearchResults();
  }
  renderTasks();
};

const hideSearchResults = () => {
  if (!elements.searchResults) return;
  elements.searchResults.replaceChildren();
  elements.searchResults.hidden = true;
};

const renderSearchResults = () => {
  if (!elements.searchResults) return;
  const term = state.searchTerm.trim().toLowerCase();
  if (!term) {
    hideSearchResults();
    return;
  }
  const matches = state.tasks
    .filter((task) => !task.deletedAt && matchesSearch(task) && matchesFilters(task))
    .sort(
      (a, b) =>
        new Date(b.updatedAt || b.createdAt || 0).getTime() -
        new Date(a.updatedAt || a.createdAt || 0).getTime()
    )
    .slice(0, 5);

  const fragment = document.createDocumentFragment();
  if (!matches.length) {
    const empty = document.createElement("p");
    empty.className = "search-suggestion__empty";
    empty.textContent = "No matches yet.";
    fragment.append(empty);
  } else {
    const list = document.createElement("ul");
    list.className = "search-suggestion__list";
    matches.forEach((task) => {
      const item = document.createElement("li");
      const button = document.createElement("button");
      button.type = "button";
      button.className = "search-suggestion__item";
      button.dataset.action = "search-select-task";
      button.dataset.taskId = task.id;
      const title = document.createElement("span");
      title.className = "search-suggestion__title";
      title.textContent = task.title;
      const meta = document.createElement("span");
      meta.className = "search-suggestion__meta";
      const metaParts = [];
      const project = getProjectById(task.projectId);
      if (project) metaParts.push(project.name);
      const company = getCompanyById(task.companyId);
      if (company && state.activeCompanyId === ALL_COMPANY_ID) {
        metaParts.push(company.name);
      }
      if (task.kind && task.kind !== "task") {
        metaParts.push(task.kind.charAt(0).toUpperCase() + task.kind.slice(1));
      }
      meta.textContent = metaParts.join(" ? ");
      button.append(title, meta);
      item.append(button);
      list.append(item);
    });
    fragment.append(list);
  }

  elements.searchResults.replaceChildren(fragment);
  elements.searchResults.hidden = false;
};

const tasksForCurrentView = () =>
  state.tasks.filter(
    (task) => !task.deletedAt && matchesActiveView(task) && matchesSearch(task) && matchesFilters(task),
  );

const openTasks = (tasks) => tasks.filter((task) => !task.completed);
const completedTasks = (tasks) => tasks.filter((task) => task.completed);

const renderCompanyTabs = () => {
  if (!elements.companyTabs) return;
  const fragment = document.createDocumentFragment();
  const makeTabButton = (companyId, label) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "toggle-pill";
    button.dataset.companyTab = companyId;
    const isActive =
      companyId === (state.activeCompanyId || DEFAULT_COMPANY.id) ||
      (!state.activeCompanyId && companyId === DEFAULT_COMPANY.id);
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
    button.textContent = label;
    return button;
  };

  const allWrapper = document.createElement("div");
  allWrapper.className = "company-tab";
  const allButton = makeTabButton(ALL_COMPANY_ID, "All");
  const allActive = state.activeCompanyId === ALL_COMPANY_ID;
  allButton.classList.toggle("active", allActive);
  allButton.setAttribute("aria-pressed", String(allActive));
  allWrapper.append(allButton);
  fragment.append(allWrapper);

  state.companies.forEach((company) => {
    const wrapper = document.createElement("div");
    wrapper.className = "company-tab";
    const select = makeTabButton(company.id, company.name);
    select.classList.toggle("active", company.id === state.activeCompanyId);
    select.setAttribute("aria-pressed", String(company.id === state.activeCompanyId));
    wrapper.append(select);

    const menuButton = document.createElement("button");
    menuButton.type = "button";
    menuButton.className = "company-tab__menu-button";
    menuButton.dataset.action = "open-company-dialog";
    menuButton.dataset.companyId = company.id;
    menuButton.setAttribute("aria-label", `Edit ${company.name}`);
    menuButton.innerHTML =
      '<svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true"><circle cx="3" cy="8" r="1.25" fill="currentColor"/><circle cx="8" cy="8" r="1.25" fill="currentColor"/><circle cx="13" cy="8" r="1.25" fill="currentColor"/></svg>';
    wrapper.append(menuButton);
    fragment.append(wrapper);
  });

  elements.companyTabs.replaceChildren(fragment);
  if (elements.companyTabsEmpty) {
    elements.companyTabsEmpty.hidden = state.companies.length !== 0;
  }
};

const renderProjectDropdown = () => {
  if (!elements.projectDropdownList) return;
  if (elements.projectDropdownToggle) {
    elements.projectDropdownToggle.disabled = state.activeCompanyId === ALL_COMPANY_ID;
  }
  if (state.activeCompanyId === ALL_COMPANY_ID) {
    elements.projectDropdownList.replaceChildren();
    if (elements.projectDropdownLabel) {
      elements.projectDropdownLabel.textContent = "All projects";
    }
    if (elements.projectDropdownMenu) {
      elements.projectDropdownMenu.setAttribute("hidden", "");
    }
    return;
  }
  const fragment = document.createDocumentFragment();
  const scopedTasks = tasksForCompany();
  const projects = getProjectsForCompany(state.activeCompanyId);
  let highlightedProjectId = "";
  let highlightedProject = null;
  if (state.activeView.type === "project") {
    const activeProject = getProjectById(state.activeView.value);
    if (activeProject && activeProject.companyId === state.activeCompanyId && !activeProject.isDefault) {
      highlightedProjectId = activeProject.id;
      highlightedProject = activeProject;
    }
  }
  if (!highlightedProjectId) {
    const rememberedId = state.companyRecents[state.activeCompanyId];
    const rememberedProject = rememberedId ? getProjectById(rememberedId) : null;
    if (rememberedProject && rememberedProject.companyId === state.activeCompanyId && !rememberedProject.isDefault) {
      highlightedProjectId = rememberedProject.id;
      highlightedProject = rememberedProject;
    }
  }
  if (!projects.length) {
    const empty = document.createElement("li");
    empty.className = "text-sm text-slate-500 px-2 py-1.5";
    empty.textContent = "Create a project to start planning.";
    fragment.append(empty);
  } else {
    projects.forEach((project) => {
      const item = document.createElement("li");
      item.className = "selector-item";
      const select = document.createElement("button");
      select.type = "button";
      select.dataset.select = "project";
      select.dataset.projectId = project.id;
      select.setAttribute("role", "option");
      const isActive = project.id === highlightedProjectId;
      select.classList.toggle("active", isActive);
      select.setAttribute("aria-selected", String(isActive));
      select.textContent = project.name;
      const meta = document.createElement("small");
      const openTasks = scopedTasks.filter(
        (task) => task.projectId === project.id && !task.completed,
      ).length;
      meta.textContent = openTasks
        ? `${openTasks} open task${openTasks === 1 ? "" : "s"}`
        : "No open tasks";
      select.append(meta);

      const actions = document.createElement("div");
      actions.className = "selector-item-actions";
      const meetingAction = document.createElement("button");
      meetingAction.type = "button";
      meetingAction.className = "selector-action-btn";
      meetingAction.dataset.action = "quick-meeting";
      meetingAction.dataset.projectId = project.id;
      meetingAction.textContent = "Meeting";
      const emailAction = document.createElement("button");
      emailAction.type = "button";
      emailAction.className = "selector-action-btn";
      emailAction.dataset.action = "quick-email";
      emailAction.dataset.projectId = project.id;
      emailAction.textContent = "Email";
      const edit = document.createElement("button");
      edit.type = "button";
      edit.className = "selector-action-btn";
      edit.dataset.action = "edit-project";
      edit.dataset.projectId = project.id;
      edit.setAttribute("aria-label", `Rename ${project.name}`);
      edit.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><path d="M4 20h4l10-10-4-4L4 16v4Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><path d="m14 6 4 4" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/></svg>';
      actions.append(meetingAction, emailAction, edit);
      item.append(select, actions);
      fragment.append(item);
    });
  }

  elements.projectDropdownList.replaceChildren(fragment);
  if (elements.projectDropdownLabel) {
    let labelText = "Select project";
    if (highlightedProject) {
      labelText = highlightedProject.name;
    }
    if (labelText === "Select project" && projects[0]) {
      labelText = projects[0].name;
    }
    elements.projectDropdownLabel.textContent = labelText;
  }
  if (elements.projectDropdownToggle) {
    const hasCompany = Boolean(state.activeCompanyId && getCompanyById(state.activeCompanyId));
    elements.projectDropdownToggle.disabled = !hasCompany;
  }
};

const updateViewCounts = () => {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const scopedTasks = tasksForCompany();

  const counts = {
    workspace: scopedTasks.filter((task) => task.projectId === "inbox" && !task.completed).length,
    today: scopedTasks.filter((task) => {
      if (!task.dueDate || task.completed) return false;
      const due = new Date(task.dueDate);
      due.setHours(0, 0, 0, 0);
      return due.getTime() === today.getTime();
    }).length,
    upcoming: scopedTasks.filter((task) => {
      if (!task.dueDate || task.completed) return false;
      const due = new Date(task.dueDate);
      due.setHours(0, 0, 0, 0);
      return due.getTime() > today.getTime();
    }).length,
  };

  if (elements.viewCounts.workspace) {
    elements.viewCounts.workspace.textContent = counts.workspace;
  }
  if (elements.viewCounts.today) {
    elements.viewCounts.today.textContent = counts.today;
  }
  if (elements.viewCounts.upcoming) {
    elements.viewCounts.upcoming.textContent = counts.upcoming;
  }
};

const populateProjectOptions = () => {
  const selects = [
    elements.quickAddProject,
    elements.dialogForm?.elements.project,
  ].filter(Boolean);

  const options = state.projects.map((project) => {
    const option = document.createElement("option");
    option.value = project.id;
    const company = getCompanyById(project.companyId);
    option.textContent = company && !company.isDefault ? `${project.name} - ${company.name}` : project.name;
    return option;
  });

  selects.forEach((select) => {
    const current = select.value;
    select.replaceChildren(...options.map((opt) => opt.cloneNode(true)));
    if (options.some((opt) => opt.value === current)) {
      select.value = current;
    } else {
      select.value = "inbox";
    }
  });
};

const populateProjectCompanyOptions = (select, selectedId) => {
  if (!select) return;
  const fragment = document.createDocumentFragment();
  state.companies.forEach((company) => {
    const option = document.createElement("option");
    option.value = company.id;
    option.textContent = company.name;
    option.disabled = Boolean(company.isDefault && state.companies.length === 1);
    fragment.append(option);
  });
  select.replaceChildren(fragment);
  select.value = selectedId && getCompanyById(selectedId) ? selectedId : DEFAULT_COMPANY.id;
};

const populateSectionOptions = (select, projectId, preferredValue) => {
  if (!select) return;
  const sections = getSectionsForProject(projectId);
  const options = sections.map((section) => {
    const option = document.createElement("option");
    option.value = section.id;
    option.textContent = section.name;
    return option;
  });
  select.replaceChildren(...options);
  const fallback = sections[0]?.id ?? "";
  select.value = options.some((opt) => opt.value === preferredValue) ? preferredValue : fallback;
};

const makeUnassignedOption = (label) => {
  const option = document.createElement("option");
  option.value = "";
  option.textContent = label;
  return option;
};

const populateDepartmentOptions = (select, preferredValue) => {
  if (!select) return;
  const options = [
    makeUnassignedOption("No department"),
    ...state.departments.map((department) => {
      const option = document.createElement("option");
      option.value = department.id;
      option.textContent = department.name;
      option.disabled = Boolean(department.isDefault && state.departments.length === 1);
      return option;
    }),
  ];
  select.replaceChildren(...options);
  select.value = options.some((opt) => opt.value === preferredValue) ? preferredValue ?? "" : "";
};

const formatMemberLabel = (member) => {
  if (!member) return "";
  const department = getDepartmentById(member.departmentId);
  return department ? `${member.name} (${department.name})` : member.name;
};

const populateMemberOptions = (select, preferredValue, departmentFilter = "") => {
  if (!select) return;
  const filteredMembers =
    departmentFilter && departmentFilter !== ""
      ? state.members.filter((member) => member.departmentId === departmentFilter)
      : state.members;

  const options = [
    makeUnassignedOption("Unassigned"),
    ...filteredMembers.map((member) => {
      const option = document.createElement("option");
      option.value = member.id;
      option.textContent = formatMemberLabel(member);
      return option;
    }),
  ];
  select.replaceChildren(...options);
  select.value = options.some((opt) => opt.value === preferredValue) ? preferredValue ?? "" : "";
};

const updateTeamSelects = () => {
  populateDepartmentOptions(elements.quickAddDepartment, elements.quickAddDepartment?.value);
  populateMemberOptions(
    elements.quickAddAssignee,
    elements.quickAddAssignee?.value,
    elements.quickAddDepartment?.value ?? ""
  );

  if (elements.dialogForm) {
    populateDepartmentOptions(
      elements.dialogForm.elements.department,
      elements.dialogForm.elements.department?.value
    );
    populateMemberOptions(
      elements.dialogForm.elements.assignee,
      elements.dialogForm.elements.assignee?.value,
      elements.dialogForm.elements.department?.value ?? ""
    );
  }

  if (elements.meetingDepartment && elements.meetingAssignee) {
    populateDepartmentOptions(elements.meetingDepartment, elements.meetingDepartment.value);
    populateMemberOptions(
      elements.meetingAssignee,
      elements.meetingAssignee.value,
      elements.meetingDepartment.value || ""
    );
  }

  if (elements.emailDepartment && elements.emailAssignee) {
    populateDepartmentOptions(elements.emailDepartment, elements.emailDepartment.value);
    populateMemberOptions(
      elements.emailAssignee,
      elements.emailAssignee.value,
      elements.emailDepartment.value || ""
    );
  }

  const memberDepartmentSelect = elements.membersForm?.elements.memberDepartment;
  populateDepartmentOptions(memberDepartmentSelect, memberDepartmentSelect?.value ?? DEFAULT_DEPARTMENT.id);
};

const updateSectionSelects = () => {
  const currentProjectQuickAdd = elements.quickAddProject?.value ?? "inbox";
  populateSectionOptions(elements.quickAddSection, currentProjectQuickAdd, elements.quickAddSection?.value);

  if (elements.dialogForm) {
    const projectId = elements.dialogForm.elements.project.value ?? "inbox";
    populateSectionOptions(
      elements.dialogForm.elements.section,
      projectId,
      elements.dialogForm.elements.section?.value
    );
  }
};

const updateActiveNav = () => {
  if (!elements.viewList) return;
  elements.viewList
    .querySelectorAll("button[data-view]")
    .forEach((button) => {
      const isActive = state.activeView.type === "view" && button.dataset.view === state.activeView.value;
      button.classList.toggle("active", isActive);
      button.setAttribute("aria-pressed", String(isActive));
    });
};

const describeView = () => {
  if (state.activeCompanyId === ALL_COMPANY_ID) {
    return { title: "All Activity", subtitle: "Latest updates across all companies." };
  }
  const { type, value } = state.activeView;
  if (type === "view") {
    if (value === "workspace") {
      return { title: "Workspace", subtitle: "All unscheduled tasks live here." };
    }
    if (value === "today") {
      const formatted = new Intl.DateTimeFormat(undefined, {
        weekday: "long",
        month: "short",
        day: "numeric",
      }).format(new Date());
      return { title: "Due Today", subtitle: formatted };
    }
    if (value === "upcoming") {
      return { title: "Upcoming 7 Days", subtitle: "Schedule ahead for the next week." };
    }
  }
  if (type === "project") {
    const project = getProjectById(value);
    if (project) {
      const company = getCompanyById(project.companyId);
      return {
        title: project.name,
        subtitle: company ? `${company.name} - Project board` : "View tasks scoped to this project.",
      };
    }
  }
  return { title: "Tasks", subtitle: "Stay organised and on track." };
};

const updateSectionActions = () => {
  if (!elements.sectionActions) return;
  const isProjectView = state.activeView.type === "project";
  elements.sectionActions.hidden = !isProjectView;
};

const updateViewToggleButtons = () => {
  elements.viewToggleButtons.forEach((button) => {
    const mode = button.dataset.viewMode;
    const isActive = mode === state.viewMode;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
  });
};

const applyViewVisibility = () => {
  const boardVisible = state.viewMode === "board";
  if (elements.boardView) elements.boardView.hidden = !boardVisible;
  if (elements.listView) elements.listView.hidden = boardVisible;
};

const renderHeader = () => {
  const { title, subtitle } = describeView();
  elements.viewTitle.textContent = title;
  elements.viewSubtitle.textContent = subtitle;
  updateSectionActions();
  updateViewToggleButtons();
};

const shouldShowProjectChip = (task) => {
  if (task.projectId === "inbox") return false;
  if (state.activeView.type === "project" && state.activeView.value === task.projectId) {
    return false;
  }
  return true;
};

const pushMetaChip = (container, text, className = "meta-chip") => {
  if (!text) return;
  const chip = document.createElement("span");
  chip.className = className;
  chip.textContent = text;
  container.append(chip);
};

const renderTaskItem = (task) => {
  const fragment = elements.taskTemplate.content.cloneNode(true);
  const item = fragment.querySelector(".task-item");
  const titleEl = fragment.querySelector(".task-title");
  const previewEl = fragment.querySelector(".task-preview");
  const assigneeEl = fragment.querySelector(".task-meta-assignee");
  const actionEl = fragment.querySelector(".task-meta-actions");
  const dueEl = fragment.querySelector(".task-meta-due");
  const editBtn = fragment.querySelector('[data-action="edit"]');

  item.dataset.taskId = task.id;
  item.dataset.priority = task.priority ?? "medium";
  titleEl.textContent = task.title;
  titleEl.title = task.title;
  previewEl.textContent = describeTaskPreview(task);
  assigneeEl.textContent = describeTaskAssignee(task);
  actionEl.textContent = describeTaskActionItems(task);
  dueEl.textContent = describeTaskDueLabel(task);

  item.classList.toggle("completed", Boolean(task.completed));
  item.classList.toggle("deleted", Boolean(task.deletedAt));

  if (editBtn) {
    editBtn.textContent = task.completed ? "View" : "Open";
    editBtn.disabled = Boolean(task.deletedAt);
  }

  return item;
};

const renderTaskTabs = (tabs, activeTab) => {
  if (!elements.taskTabList) return;
  if (tabs.length <= 1) {
    elements.taskTabList.replaceChildren();
    elements.taskTabList.hidden = true;
    return;
  }

  const fragment = document.createDocumentFragment();
  tabs.forEach(({ id, label }) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "toggle-pill";
    button.dataset.taskTab = id;
    const isActive = id === activeTab;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
    button.textContent = label;
    fragment.append(button);
  });

  elements.taskTabList.hidden = false;
  elements.taskTabList.replaceChildren(fragment);
};

const createEmptyState = (message) => {
  const empty = document.createElement("p");
  empty.className = "empty-state";
  empty.textContent = message;
  return empty;
};

const getDeletedTasksForProject = (projectId) =>
  state.tasks.filter(
    (task) => task.projectId === projectId && task.deletedAt && matchesSearch(task)
  );

const setProjectGroupLimit = (projectId, groupId, value) => {
  if (!projectId || !groupId) return;
  if (!state.projectSectionLimits[projectId]) {
    state.projectSectionLimits[projectId] = {};
  }
  state.projectSectionLimits[projectId][groupId] = value;
};

const resolveProjectGroupLimit = (projectId, groupId, total) => {
  const projectLimits = state.projectSectionLimits[projectId] ?? {};
  const raw = projectLimits[groupId];
  if (Number.isFinite(raw) && raw > 0) {
    return Math.min(total, raw);
  }
  if (total > 5) return 5;
  return total;
};

const renderProjectOverview = (projectId, tasks) => {
  const container = document.createElement("div");
  container.className = "project-overview";

  PROJECT_OVERVIEW_GROUPS.forEach((group) => {
    const groupTasks = tasks
      .filter(group.filter)
      .sort(
        (a, b) =>
          new Date(b.updatedAt || b.createdAt || 0).getTime() -
          new Date(a.updatedAt || a.createdAt || 0).getTime()
      );

    const total = groupTasks.length;
    const limit = total ? resolveProjectGroupLimit(projectId, group.id, total) : 0;
    const visibleTasks = limit ? groupTasks.slice(0, limit) : [];

    const card = document.createElement("article");
    card.className = "overview-section";

    const header = document.createElement("header");
    header.className = "overview-section__header";
    const title = document.createElement("h4");
    title.textContent = group.title;
    header.append(title);

    if (total > visibleTasks.length) {
      const moreButton = document.createElement("button");
      moreButton.type = "button";
      moreButton.className = "see-more";
      moreButton.dataset.action = "project-overview-expand";
      moreButton.dataset.projectId = projectId;
      moreButton.dataset.groupId = group.id;
      moreButton.dataset.total = String(total);
      moreButton.textContent = "See more";
      header.append(moreButton);
    }

    card.append(header);

    const body = document.createElement("div");
    body.className = "overview-section__body resizable-block";

    if (visibleTasks.length) {
      const list = document.createElement("ul");
      list.className = "task-list compact";
      visibleTasks.forEach((task) => list.append(renderTaskItem(task)));
      body.append(list);
    } else {
      const empty = document.createElement("p");
      empty.className = "overview-section__empty";
      empty.textContent = `No ${group.title.toLowerCase()}.`;
      body.append(empty);
    }

    card.append(body);
    container.append(card);
  });

  if (!container.childElementCount) {
    const empty = document.createElement("p");
    empty.className = "empty-state";
    empty.textContent = "No tasks recorded for this project yet.";
    container.append(empty);
  }

  return container;
};

const expandProjectOverviewSection = (projectId, groupId, total) => {
  if (!projectId || !groupId) return;
  const numericTotal = Number.parseInt(total ?? "0", 10);
  const nextLimit = Number.isFinite(numericTotal) && numericTotal > 0 ? numericTotal : 50;
  setProjectGroupLimit(projectId, groupId, nextLimit);
  renderTasks();
};

const handleFilterChange = (event) => {
  const select = event.target.closest("select[data-filter]");
  if (!select || !state.filters) return;
  const key = select.dataset.filter;
  if (!key) return;
  const value = select.value || "";
  if (state.filters[key] === value) return;
  state.filters = { ...state.filters, [key]: value };
  renderFilterBar();
  renderTasks();
};

const handleFilterBarClick = (event) => {
  const button = event.target.closest('[data-action="clear-filters"]');
  if (!button || !state.filters) return;
  event.preventDefault();
  if (!isFilterActive()) return;
  state.filters = { member: "", department: "", priority: "", due: "" };
  renderFilterBar();
  renderTasks();
};

const renderProjectSectionsList = (
  projectId,
  tasks,
  {
    includeEmpty = true,
    emptyMessage = "No tasks recorded for this project yet.",
    emptySectionMessage = "No tasks in this section.",
    sorter = compareTasks,
  } = {}
) => {
  const container = document.createElement("div");
  container.className = "list-sections space-y-4";
  const sections = getSectionsForProject(projectId);

  if (!sections.length) {
    container.append(createEmptyState(emptyMessage));
    return container;
  }

  sections.forEach((section) => {
    const sectionTasks = tasks.filter((task) => task.sectionId === section.id);
    if (!includeEmpty && sectionTasks.length === 0) {
      return;
    }

    const card = document.createElement("article");
    card.className = "list-section";
    card.dataset.sectionId = section.id;

    const header = document.createElement("header");
    header.className = "list-section__header";

    const title = document.createElement("h4");
    title.textContent = section.name;
    header.append(title);

    const count = document.createElement("span");
    count.className = "list-section__count";
    count.textContent = String(sectionTasks.length);
    header.append(count);

    const actions = document.createElement("div");
    actions.className = "list-section__actions";
    const edit = document.createElement("button");
    edit.type = "button";
    edit.className = "ghost-button small";
    edit.dataset.action = "edit-section";
    edit.dataset.sectionId = section.id;
    edit.textContent = "Edit";
    actions.append(edit);
    header.append(actions);
    card.append(header);

    const list = document.createElement("ul");
    list.className = "task-list";
    if (sectionTasks.length) {
      sectionTasks.sort(sorter).forEach((task) => list.append(renderTaskItem(task)));
    }
    card.append(list);

    if (!sectionTasks.length && includeEmpty) {
      const placeholder = document.createElement("p");
      placeholder.className = "list-section__empty";
      placeholder.textContent = emptySectionMessage;
      card.append(placeholder);
    }

    container.append(card);
  });

  if (!container.childElementCount) {
    container.append(createEmptyState(emptyMessage));
  }

  return container;
};

const renderSimpleTaskList = (tasks, emptyMessage) => {
  const container = document.createElement("div");
  if (tasks.length) {
    const list = document.createElement("ul");
    list.className = "task-list";
    tasks.forEach((task) => list.append(renderTaskItem(task)));
    container.append(list);
    return container;
  }
  container.append(createEmptyState(emptyMessage));
  return container;
};

const renderAllTaskOverview = () => {
  if (!elements.taskTabPanels || !elements.taskTabList) return;
  const tabs = [
    { id: "created", label: "Created" },
    { id: "updated", label: "Updated" },
    { id: "completed", label: "Completed" },
    { id: "deleted", label: "Deleted" },
  ];
  if (!tabs.some((tab) => tab.id === state.activeAllTab)) {
    state.activeAllTab = "created";
  }
  renderTaskTabs(tabs, state.activeAllTab);

  const panel = document.createElement("div");
  panel.className = "task-pane";

  let scopedTasks = [];
  if (state.activeAllTab === "created") {
    scopedTasks = state.tasks
      .filter((task) => !task.deletedAt && matchesSearch(task) && matchesFilters(task))
      .sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
  } else if (state.activeAllTab === "updated") {
    scopedTasks = state.tasks
      .filter((task) => !task.deletedAt && matchesSearch(task) && matchesFilters(task))
      .sort(
        (a, b) =>
          new Date(b.updatedAt || b.createdAt).getTime() -
          new Date(a.updatedAt || a.createdAt).getTime()
      );
  } else if (state.activeAllTab === "deleted") {
    scopedTasks = state.tasks
      .filter((task) => task.deletedAt && matchesSearch(task) && matchesFilters(task))
      .sort((a, b) => new Date(b.deletedAt ?? 0).getTime() - new Date(a.deletedAt ?? 0).getTime());
  } else {
    scopedTasks = state.tasks
      .filter((task) => task.completed && matchesSearch(task) && matchesFilters(task))
      .sort((a, b) => {
        const aTime = new Date(a.completedAt || a.updatedAt || a.createdAt || 0).getTime();
        const bTime = new Date(b.completedAt || b.updatedAt || b.createdAt || 0).getTime();
        return bTime - aTime;
      });
  }

  panel.append(
    renderSimpleTaskList(
      scopedTasks,
      state.activeAllTab === "deleted"
        ? "No deleted tasks yet."
        : state.activeAllTab === "completed"
          ? "No completed tasks yet."
          : "No tasks match this filter right now."
    )
  );
  elements.taskTabPanels.replaceChildren(panel);
};

const renderListView = (tasks) => {
  if (!elements.taskTabPanels || !elements.taskTabList) return;

  if (state.activeCompanyId === ALL_COMPANY_ID) {
    renderAllTaskOverview();
    return;
  }

  const isProjectScope = state.activeView.type === "project";
  const tabs = isProjectScope
    ? [
        { id: "active", label: "Active" },
        { id: "completed", label: "Completed" },
        { id: "deleted", label: "Deleted" },
      ]
    : [{ id: "active", label: "Active" }];

  if (!tabs.some((tab) => tab.id === state.activeTaskTab)) {
    state.activeTaskTab = "active";
  }

  renderTaskTabs(tabs, state.activeTaskTab);

  const panel = document.createElement("div");
  panel.className = "task-pane";

  if (state.activeTaskTab === "active") {
    const activeTasks = openTasks(tasks);
    if (isProjectScope) {
      panel.append(renderProjectOverview(state.activeView.value, activeTasks));
    } else {
      panel.append(
        renderSimpleTaskList(activeTasks.sort(compareTasks), "No tasks here yet. Add one to get started.")
      );
    }
  } else if (state.activeTaskTab === "completed") {
    const completed = completedTasks(tasks).sort(compareTasks);
    panel.append(
      renderProjectSectionsList(state.activeView.value, completed, {
        includeEmpty: false,
        emptyMessage: "No completed tasks yet.",
      })
    );
  } else if (state.activeTaskTab === "deleted") {
    const deleted = getDeletedTasksForProject(state.activeView.value).sort(
      (a, b) => new Date(b.deletedAt ?? 0).getTime() - new Date(a.deletedAt ?? 0).getTime()
    );
    panel.append(
      renderProjectSectionsList(state.activeView.value, deleted, {
        includeEmpty: false,
        emptyMessage: "No deleted tasks archived for this project.",
        emptySectionMessage: "No deleted tasks for this section.",
        sorter: (a, b) =>
          new Date(b.deletedAt ?? 0).getTime() - new Date(a.deletedAt ?? 0).getTime(),
      })
    );
  }

  elements.taskTabPanels.replaceChildren(panel);
};
const createBoardCard = (task) => {
  const card = document.createElement("div");
  card.className = "board-card";
  card.classList.add(`priority-${task.priority ?? "medium"}`);
  card.draggable = true;
  card.dataset.taskId = task.id;

  const title = document.createElement("h4");
  title.textContent = task.title;
  card.append(title);

  const meta = document.createElement("div");
  meta.className = "meta-line";
  const due = describeDueDate(task.dueDate);
  const member = getMemberById(task.assigneeId);
  const metaParts = [];
  if (due.label) metaParts.push(due.label);
  if (member) metaParts.push(member.name);
  const priorityLabel = PRIORITY_LABELS[task.priority];
  if (priorityLabel && task.priority !== "medium") metaParts.push(priorityLabel);
  meta.textContent = metaParts.join(" | ");
  card.append(meta);

  if (Array.isArray(task.attachments) && task.attachments.length) {
    const attachmentMeta = document.createElement("div");
    attachmentMeta.className = "meta-line";
    attachmentMeta.textContent = `${task.attachments.length} attachment${task.attachments.length === 1 ? "" : "s"}`;
    card.append(attachmentMeta);
  }

  if (task.description) {
    const note = document.createElement("div");
    note.className = "meta-line";
    note.textContent = "Notes available";
    card.append(note);
  }

  card.addEventListener("dragstart", handleBoardDragStart);
  card.addEventListener("dragend", handleBoardDragEnd);
  card.addEventListener("dblclick", () => openTaskEditor(task));

  return card;
};

const createBoardColumn = (section, tasks) => {
  const column = document.createElement("article");
  column.className = "board-column drop-zone";
  column.dataset.sectionId = section.id;

  const header = document.createElement("header");
  header.dataset.sectionId = section.id;
  header.draggable = true;
  header.addEventListener("dragstart", handleSectionDragStart);
  header.addEventListener("dragover", handleSectionDragOver);
  header.addEventListener("dragleave", handleSectionDragLeave);
  header.addEventListener("drop", handleSectionDrop);
  header.addEventListener("dragend", handleSectionDragEnd);

  const title = document.createElement("h3");
  title.textContent = section.name;
  header.append(title);

  const count = document.createElement("span");
  count.className = "count";
  count.textContent = tasks.length;
  header.append(count);

  const actions = document.createElement("div");
  actions.className = "section-actions";
  const editButton = document.createElement("button");
  editButton.type = "button";
  editButton.className = "ghost-button small";
  editButton.dataset.action = "edit-section";
  editButton.dataset.sectionId = section.id;
  editButton.textContent = "Edit";
  actions.append(editButton);
  header.append(actions);

  column.append(header);

  const body = document.createElement("div");
  body.className = "board-tasks drop-zone";
  body.dataset.sectionId = section.id;
  body.addEventListener("dragover", handleBoardDragOver);
  body.addEventListener("dragleave", handleBoardDragLeave);
  body.addEventListener("drop", handleBoardDrop);

  if (tasks.length === 0) {
    body.classList.add("empty");
  } else {
    tasks.forEach((task) => body.append(createBoardCard(task)));
  }

  column.append(body);
  return column;
};

const renderBoardView = (tasks) => {
  if (state.activeView.type !== "project") {
    elements.boardColumns.replaceChildren();
    elements.boardEmptyState.hidden = false;
    elements.boardEmptyState.textContent = "Switch to a specific project to use the board view.";
    return;
  }

  const projectId = state.activeView.value;
  ensureSectionForProject(projectId);
  const sections = getSectionsForProject(projectId);

  if (!sections.length) {
    elements.boardColumns.replaceChildren();
    elements.boardEmptyState.hidden = false;
    elements.boardEmptyState.textContent = "Add a section to start organising this board.";
    return;
  }

  const fragment = document.createDocumentFragment();
  sections.forEach((section) => {
    const sectionTasks = tasks
      .filter((task) => task.sectionId === section.id && !task.completed)
      .sort(compareTasks);
    fragment.append(createBoardColumn(section, sectionTasks));
  });

  elements.boardColumns.replaceChildren(fragment);
  elements.boardEmptyState.hidden = true;
};

const renderTasks = () => {
  const tasks = tasksForCurrentView();
  if (state.activeCompanyId === ALL_COMPANY_ID && state.viewMode !== "list") {
    state.viewMode = "list";
    updateViewToggleButtons();
    applyViewVisibility();
  }
  if (state.viewMode === "board") {
    renderBoardView(tasks);
  } else {
    renderListView(tasks);
  }
};

const renderTeamStatus = () => {
  if (!elements.teamStatus) return;
  const fragment = document.createDocumentFragment();
  const scopedTasks = tasksForCompany();
  const summaries = state.departments.map((department) => {
    const members = state.members.filter((member) => member.departmentId === department.id);
    const activeTasks = scopedTasks.filter((task) => task.departmentId === department.id && !task.completed).length;
    return { department, members, activeTasks };
  }).filter((entry) => entry.department.isDefault || entry.members.length || entry.activeTasks);

  if (!summaries.length) {
    const empty = document.createElement("li");
    empty.textContent = "No departments or members yet.";
    fragment.append(empty);
  } else {
    summaries.forEach(({ department, members, activeTasks }) => {
      const item = document.createElement("li");
      const info = document.createElement("div");
      const name = document.createElement("p");
      name.className = "font-medium text-slate-800";
      name.textContent = department.name;
      const meta = document.createElement("p");
      meta.className = "text-xs text-slate-500";
      meta.textContent = `${members.length} member${members.length === 1 ? "" : "s"}`;
      info.append(name, meta);

      const chip = document.createElement("span");
      chip.className = "chip";
      chip.textContent = `${activeTasks} active`;

      item.append(info, chip);
      fragment.append(item);
    });
  }

  elements.teamStatus.replaceChildren(fragment);
};

const formatRelativeTime = (isoString) => {
  if (!isoString) return "Recently";
  const date = new Date(isoString);
  if (Number.isNaN(date.getTime())) return "Recently";
  const diffMs = Date.now() - date.getTime();
  const diffMinutes = Math.floor(diffMs / (1000 * 60));
  if (diffMinutes < 1) return "Just now";
  if (diffMinutes < 60) return `${diffMinutes}m ago`;
  const diffHours = Math.floor(diffMinutes / 60);
  if (diffHours < 24) return `${diffHours}h ago`;
  const diffDays = Math.floor(diffHours / 24);
  if (diffDays < 7) return `${diffDays}d ago`;
  const diffWeeks = Math.floor(diffDays / 7);
  if (diffWeeks < 4) return `${diffWeeks}w ago`;
  const diffMonths = Math.floor(diffDays / 30);
  if (diffMonths < 12) return `${diffMonths}mo ago`;
  const diffYears = Math.floor(diffDays / 365);
  return `${diffYears}y ago`;
};

const renderActivityFeed = () => {
  if (!elements.activityFeed) return;
  const fragment = document.createDocumentFragment();
  const limit = Number.isFinite(state.recentActivityLimit) && state.recentActivityLimit > 0 ? state.recentActivityLimit : 10;
  const sorted = [...tasksForCompany()].sort((a, b) => {
    const bTime = new Date(b.updatedAt || b.createdAt || 0).getTime();
    const aTime = new Date(a.updatedAt || a.createdAt || 0).getTime();
    return bTime - aTime;
  });
  const recent = sorted.slice(0, limit);

  if (!recent.length) {
    const empty = document.createElement("p");
    empty.className = "text-sm text-slate-500";
    empty.textContent = "No activity yet. Add your first task.";
    fragment.append(empty);
  } else {
    recent.forEach((task) => {
      const item = document.createElement("button");
      item.type = "button";
      item.className = "activity-item bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-left";
      item.dataset.taskLink = task.id;
      item.setAttribute("aria-label", `Open task "${task.title}"`);
      if (task.deletedAt) {
        item.classList.add("activity-item--deleted");
      }

      const title = document.createElement("span");
      title.className = "text-sm font-medium text-slate-800";
      title.textContent = task.title;

      const meta = document.createElement("p");
      meta.className = "text-xs text-slate-500 flex items-center flex-wrap gap-2";
      const timeBadge = document.createElement("span");
      timeBadge.textContent = formatRelativeTime(task.updatedAt || task.createdAt);
      meta.append(timeBadge);

      const project = getProjectById(task.projectId);
      if (project) {
        const projectChip = document.createElement("span");
        projectChip.className = "chip";
        projectChip.textContent = project.name;
        meta.append(projectChip);
      }

      const company = getCompanyById(task.companyId);
      if (company && state.activeCompanyId === ALL_COMPANY_ID) {
        const companyChip = document.createElement("span");
        companyChip.className = "chip subtle";
        companyChip.textContent = company.name;
        meta.append(companyChip);
      }

      const member = getMemberById(task.assigneeId);
      if (member) {
        const memberChip = document.createElement("span");
        memberChip.className = "chip";
        memberChip.textContent = member.name;
        meta.append(memberChip);
      }

      item.append(title, meta);
      fragment.append(item);
    });
  }

  if (sorted.length > recent.length) {
    const moreWrapper = document.createElement("div");
    moreWrapper.className = "activity-more";
    const moreButton = document.createElement("button");
    moreButton.type = "button";
    moreButton.className = "see-more";
    moreButton.dataset.action = "show-more-activity";
    moreButton.textContent = "Show more";
    moreWrapper.append(moreButton);
    fragment.append(moreWrapper);
  }

  elements.activityFeed.replaceChildren(fragment);
};

const buildUserguideFragment = (entries) => {
  const fragment = document.createDocumentFragment();
  let currentList = null;
  let listType = "";

  const closeList = () => {
    if (currentList) {
      fragment.append(currentList);
      currentList = null;
      listType = "";
    }
  };

  entries.forEach((rawLine) => {
    const line = rawLine.trim();
    if (!line) return;
    if (line.startsWith("### ")) {
      closeList();
      const heading = document.createElement("h3");
      heading.textContent = line.slice(4).trim();
      fragment.append(heading);
      return;
    }
    if (line.startsWith("## ")) {
      closeList();
      const heading = document.createElement("h2");
      heading.textContent = line.slice(3).trim();
      fragment.append(heading);
      return;
    }
    if (/^[-]\s+/.test(line)) {
      if (listType !== "ul") {
        closeList();
        currentList = document.createElement("ul");
        currentList.className = "userguide-list__unordered";
        listType = "ul";
      }
      const item = document.createElement("li");
      item.textContent = line.replace(/^[-]\s+/, '').trim();
      currentList.append(item);
      return;
    }
    if (/^\d+\./.test(line)) {
      if (listType !== "ol") {
        closeList();
        currentList = document.createElement("ol");
        currentList.className = "userguide-list__ordered";
        listType = "ol";
      }
      const item = document.createElement("li");
      item.textContent = line.replace(/^\d+\.\s*/, '').trim();
      currentList.append(item);
      return;
    }
    closeList();
    const paragraph = document.createElement("p");
    paragraph.textContent = line;
    fragment.append(paragraph);
  });

  closeList();
  return fragment;
};

const renderUserguidePanel = () => {
  if (!elements.userguidePanel) return;
  if (elements.userguideList) {
    if (state.isEditingUserguide) {
      elements.userguideList.hidden = true;
    } else {
      const fragment = buildUserguideFragment(state.userguide);
      elements.userguideList.replaceChildren(fragment);
      elements.userguideList.hidden = false;
    }
  }
  if (elements.userguideForm) {
    elements.userguideForm.hidden = !state.isEditingUserguide;
  }
  if (elements.userguideEditToggle) {
    elements.userguideEditToggle.hidden = state.isEditingUserguide;
  }
  if (state.isEditingUserguide && elements.userguideEditor) {
    elements.userguideEditor.value = state.userguide.join("\n");
  }
};

const updateDashboardMetrics = () => {
  const filter = state.metricsFilter || 'all';
  const scopedTasks = tasksForCompany();
  let value = 0;

  if (filter === 'completed') {
    value = scopedTasks.filter((task) => task.completed).length;
  } else if (filter === 'all') {
    value = scopedTasks.filter((task) => !task.completed).length;
  } else {
    value = scopedTasks.filter((task) => !task.completed && task.priority === filter).length;
  }

  if (elements.metricFilter && elements.metricFilter.value !== filter) {
    elements.metricFilter.value = filter;
  }

  if (elements.activeTasksMetric) {
    elements.activeTasksMetric.textContent = value;
  }
};

const applySettings = () => {
  state.settings = normaliseSettings(state.settings);
  const { profile, theme, display } = state.settings;
  if (elements.profileNameDisplay) {
    elements.profileNameDisplay.textContent = profile.name;
  }
  if (elements.profileAvatar && profile.photo) {
    elements.profileAvatar.src = profile.photo;
  }
  const root = document.documentElement;
  if (theme?.accent) {
    root.style.setProperty('--accent', theme.accent);
    root.style.setProperty('--accent-soft', hexToRgba(theme.accent, 0.15));
    root.style.setProperty('--accent-strong', hexToRgba(theme.accent, 0.25));
  }
  const priorities = theme?.priorities ?? {};
  if (priorities.critical) root.style.setProperty('--priority-critical', priorities.critical);
  if (priorities.veryHigh) root.style.setProperty('--priority-very-high', priorities.veryHigh);
  if (priorities.high) root.style.setProperty('--priority-high', priorities.high);
  if (priorities.medium) root.style.setProperty('--priority-medium', priorities.medium);
  if (priorities.low) root.style.setProperty('--priority-low', priorities.low);
  if (priorities.optional) root.style.setProperty('--priority-optional', priorities.optional);
  const fontScale = clampFontScale(display?.fontScale || 1);
  root.style.setProperty('--workspace-font-scale', fontScale.toString());
};

const renderSidebar = () => {
  renderCompanyTabs();
  renderProjectDropdown();
  updateActiveNav();
  updateViewCounts();
  renderActivityFeed();
};

const renderFilterBar = () => {
  if (!elements.filterBar || !state.filters) return;
  const updateFilterValue = (key, value) => {
    state.filters = { ...state.filters, [key]: value };
  };

  if (elements.filterMember) {
    const fragment = document.createDocumentFragment();
    const allOption = document.createElement("option");
    allOption.value = "";
    allOption.textContent = "All members";
    fragment.append(allOption);

    const unassignedOption = document.createElement("option");
    unassignedOption.value = FILTER_MEMBER_UNASSIGNED;
    unassignedOption.textContent = "Unassigned";
    fragment.append(unassignedOption);

    [...state.members]
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }))
      .forEach((member) => {
        const option = document.createElement("option");
        option.value = member.id;
        option.textContent = member.name;
        fragment.append(option);
      });

    elements.filterMember.replaceChildren(fragment);
    let memberValue = state.filters.member || "";
    if (
      memberValue &&
      memberValue !== FILTER_MEMBER_UNASSIGNED &&
      !state.members.some((member) => member.id === memberValue)
    ) {
      memberValue = "";
      updateFilterValue("member", "");
    }
    elements.filterMember.value = memberValue;
  }

  if (elements.filterDepartment) {
    const fragment = document.createDocumentFragment();
    const allOption = document.createElement("option");
    allOption.value = "";
    allOption.textContent = "All departments";
    fragment.append(allOption);

    [...state.departments]
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }))
      .forEach((department) => {
        const option = document.createElement("option");
        option.value = department.id;
        option.textContent = department.name;
        fragment.append(option);
      });

    elements.filterDepartment.replaceChildren(fragment);
    let departmentValue = state.filters.department || "";
    if (departmentValue && !state.departments.some((dept) => dept.id === departmentValue)) {
      departmentValue = "";
      updateFilterValue("department", "");
    }
    elements.filterDepartment.value = departmentValue;
  }

  if (elements.filterPriority) {
    const allowedPriorities = new Set([
      "",
      "critical",
      "very-high",
      "high",
      "medium",
      "low",
      "optional",
    ]);
    if (!allowedPriorities.has(state.filters.priority || "")) {
      updateFilterValue("priority", "");
    }
    elements.filterPriority.value = state.filters.priority || "";
  }

  if (elements.filterDue) {
    const allowedDue = new Set(["", "overdue", "today", "week", "none"]);
    if (!allowedDue.has(state.filters.due || "")) {
      updateFilterValue("due", "");
    }
    elements.filterDue.value = state.filters.due || "";
  }

  if (elements.filterClear) {
    elements.filterClear.disabled = !isFilterActive();
  }
};

const syncQuickAddSelectors = () => {
  const preferred = getPreferredProjectId();
  const defaultProjectId =
    preferred || (state.activeCompanyId === DEFAULT_COMPANY.id ? DEFAULT_PROJECT.id : "");
  if (elements.quickAddProject) {
    elements.quickAddProject.value = defaultProjectId;
  }
  populateSectionOptions(elements.quickAddSection, elements.quickAddProject?.value ?? "inbox");
};

const render = () => {
  applySettings();
  renderSidebar();
  renderFilterBar();
  renderHeader();
  updateDashboardMetrics();
  populateProjectOptions();
  updateTeamSelects();
  updateSectionSelects();
  renderUserguidePanel();
  applyViewVisibility();
  renderTasks();
  syncQuickAddSelectors();
};

const setViewMode = (mode) => {
  if (mode === state.viewMode) return;
  if (mode === "board" && state.activeView.type !== "project") {
    window.alert("Board view is available when a specific project is selected.");
    return;
  }
  state.viewMode = mode;
  savePreferences();
  updateViewToggleButtons();
  applyViewVisibility();
  renderTasks();
};

const setActiveView = (type, value) => {
  let nextType = type;
  let nextValue = value;
  if (type === "project") {
    const project = getProjectById(value);
    if (!project) {
      console.warn("Attempted to open unknown project", value);
      nextType = "view";
      nextValue = "workspace";
    } else {
      rememberProjectSelection(project.id);
    }
  }

  state.activeView = { type: nextType, value: nextValue };
  if (nextType !== "project") {
    state.activeTaskTab = "active";
  }
  if (state.viewMode === "board" && nextType !== "project") {
    state.viewMode = "list";
  }
  if (state.isUserguideOpen) {
    closeUserguide();
  }
  savePreferences();
  render();
};
const TASK_KINDS = new Set(["task", "meeting", "email"]);

const normaliseTaskKind = (value) => {
  if (typeof value !== "string") return "task";
  const trimmed = value.trim().toLowerCase();
  return TASK_KINDS.has(trimmed) ? trimmed : "task";
};

const normaliseTaskSource = (value) => {
  if (typeof value !== "string") return "manual";
  const trimmed = value.trim().toLowerCase();
  return trimmed || "manual";
};

const normaliseTaskLinks = (links) => {
  if (!Array.isArray(links)) return [];
  return links
    .map((link) => {
      const title = typeof link?.title === "string" ? link.title.trim() : "";
      const url = typeof link?.url === "string" ? link.url.trim() : "";
      return { title, url };
    })
    .filter((entry) => entry.title || entry.url);
};

const autoResizeTextarea = (textarea) => {
  if (!textarea) return;
  textarea.style.height = "auto";
  textarea.style.height = `${textarea.scrollHeight}px`;
};

const clampTextareaHeight = (value) =>
  Math.min(Math.max(value, MIN_TEXTAREA_HEIGHT), MAX_TEXTAREA_HEIGHT);

const saveTextareaHeight = (key, height) => {
  if (!key || AUTOSIZE_RESET_KEYS.has(key)) return;
  const value = clampTextareaHeight(height);
  if (!state.textareaHeights) {
    state.textareaHeights = {};
  }
  if (state.textareaHeights[key] === value) return;
  state.textareaHeights = { ...state.textareaHeights, [key]: value };
  savePreferences();
};

const applySavedTextareaHeight = (textarea) => {
  if (!textarea) return;
  const key = textarea.dataset.autosizeKey;
  const defaultHeight = clampTextareaHeight(
    Number.parseInt(textarea.dataset.autosizeDefault ?? "180", 10)
  );
  const stored = AUTOSIZE_RESET_KEYS.has(key)
    ? defaultHeight
    : clampTextareaHeight(state.textareaHeights?.[key] ?? defaultHeight);
  textarea.style.minHeight = `${defaultHeight}px`;
  textarea.style.height = `${stored}px`;
};

const registerAutosizeTextarea = (textarea) => {
  if (!textarea || textarea.dataset.autosizeReady === "true") return;
  const key = textarea.dataset.autosizeKey;
  const defaultHeight = clampTextareaHeight(
    Number.parseInt(textarea.dataset.autosizeDefault ?? "180", 10)
  );
  applySavedTextareaHeight(textarea);
  textarea.style.resize = "vertical";
  textarea.dataset.autosizeReady = "true";

  const syncHeight = () => {
    if (!key) return;
    const baseline = AUTOSIZE_RESET_KEYS.has(key)
      ? defaultHeight
      : clampTextareaHeight(state.textareaHeights?.[key] ?? defaultHeight);
    textarea.style.height = "auto";
    const target = Math.max(baseline, textarea.scrollHeight);
    textarea.style.height = `${target}px`;
    saveTextareaHeight(key, target);
  };

  const persistResize = () => {
    if (!key || AUTOSIZE_RESET_KEYS.has(key)) return;
    saveTextareaHeight(key, textarea.offsetHeight);
  };

  textarea.addEventListener("input", syncHeight);
  textarea.addEventListener("mouseup", persistResize);
  textarea.addEventListener("touchend", persistResize);
  syncHeight();
};

const initialiseTextareaAutosize = () => {
  const textareas = document.querySelectorAll("[data-autosize-key]");
  textareas.forEach((textarea) => registerAutosizeTextarea(textarea));
};

const createLinkRow = (list, link = {}) => {
  if (!list) return null;
  const row = document.createElement("li");
  row.className = "link-row";
  row.dataset.linkRow = "true";

  const titleField = document.createElement("input");
  titleField.type = "text";
  titleField.placeholder = "Title";
  titleField.className = "field-input";
  titleField.value = link.title ?? "";
  titleField.dataset.linkTitle = "true";

  const urlField = document.createElement("input");
  urlField.type = "url";
  urlField.placeholder = "https://...";
  urlField.className = "field-input";
  urlField.value = link.url ?? "";
  urlField.dataset.linkUrl = "true";

  const removeButton = document.createElement("button");
  removeButton.type = "button";
  removeButton.className = "ghost-button small";
  removeButton.dataset.action = "remove-link";
  removeButton.textContent = "Remove";

  row.append(titleField, urlField, removeButton);
  list.append(row);
  return row;
};

const resetLinkList = (list, links = []) => {
  if (!list) return;
  list.replaceChildren();
  if (!links.length) {
    createLinkRow(list);
    return;
  }
  links.forEach((link) => createLinkRow(list, link));
};

const collectLinks = (list) => {
  if (!list) return [];
  return [...list.querySelectorAll('[data-link-row]')]
    .map((row) => {
      const titleInput = row.querySelector('[data-link-title]');
      const urlInput = row.querySelector('[data-link-url]');
      const title = titleInput?.value?.trim() ?? "";
      const url = urlInput?.value?.trim() ?? "";
      return { title, url };
    })
    .filter((entry) => entry.title || entry.url);
};

const normaliseActionItem = (item) => {
  if (!item) return null;
  const title = typeof item.title === "string" ? item.title.trim() : "";
  if (!title) return null;
  const completed = Boolean(item.completed);
  const id = item.id || generateId("action-item");
  return { id, title, completed };
};

const commitMeetingActionDraft = () => {
  const committed = (state.meetingActionDraft || [])
    .map((item) => normaliseActionItem(item))
    .filter(Boolean);
  state.meetingActionDraft = committed;
  return committed;
};

const setMeetingActionDraft = (items = []) => {
  const normalised = items
    .map((item) => normaliseActionItem(item))
    .filter(Boolean);
  state.meetingActionDraft = normalised;
};

const addMeetingActionItems = (labels = []) => {
  const additions = (Array.isArray(labels) ? labels : [labels])
    .map((label) => {
      if (typeof label !== "string") return null;
      const trimmed = label.trim();
      if (!trimmed) return null;
      return { id: generateId("action-item"), title: trimmed, completed: false };
    })
    .filter(Boolean);
  if (!additions.length) return;
  state.meetingActionDraft = [...(state.meetingActionDraft || []), ...additions];
  renderMeetingActionItems();
};

const updateMeetingActionItem = (id, updates) => {
  if (!id || !updates) return;
  state.meetingActionDraft = (state.meetingActionDraft || []).map((item) => {
    if (item.id !== id) return item;
    const next = { ...item, ...updates };
    return normaliseActionItem(next) ?? item;
  });
  renderMeetingActionItems();
};

const removeMeetingActionItem = (id) => {
  if (!id) return;
  state.meetingActionDraft = (state.meetingActionDraft || []).filter((item) => item.id !== id);
  renderMeetingActionItems();
};

const renderMeetingActionItems = () => {
  const list = elements.meetingActionList;
  if (!list) return;
  list.replaceChildren();
  const items = state.meetingActionDraft || [];
  if (!items.length) {
    const empty = document.createElement("p");
    empty.className = "action-items-empty";
    empty.textContent = "No action items captured yet.";
    list.append(empty);
    return;
  }
  items.forEach((item) => {
    const row = document.createElement("div");
    row.className = "action-item-row";
    row.dataset.actionItemId = item.id;
    if (item.completed) {
      row.classList.add("completed");
    }

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.className = "action-item-checkbox";
    checkbox.checked = item.completed;
    checkbox.dataset.action = "meeting-toggle-action";
    checkbox.dataset.actionItemId = item.id;

    const textInput = document.createElement("input");
    textInput.type = "text";
    textInput.className = "field-input action-item-input";
    textInput.value = item.title;
    textInput.placeholder = "Describe the action item";
    textInput.dataset.actionItemId = item.id;

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "ghost-button small";
    removeBtn.dataset.action = "meeting-remove-action";
    removeBtn.dataset.actionItemId = item.id;
    removeBtn.textContent = "Remove";

    row.append(checkbox, textInput, removeBtn);
    list.append(row);
  });
};

const normaliseThreadMessage = (entry) => {
  if (!entry) return null;
  const source =
    typeof entry === "string"
      ? entry
      : typeof entry?.body === "string"
        ? entry.body
        : typeof entry?.content === "string"
          ? entry.content
          : typeof entry?.text === "string"
            ? entry.text
            : "";
  const body = typeof source === "string" ? source.trim() : "";
  if (!body) return null;
  const completed = Boolean(entry?.completed);
  const id = entry?.id || generateId("thread-message");
  const heightSource =
    typeof entry?.height === "number"
      ? entry.height
      : typeof entry?.height === "string"
        ? Number.parseFloat(entry.height)
        : NaN;
  const height = clampThreadHeight(heightSource);
  return { id, body, completed, height: height ?? null };
};

const commitEmailThreadDraft = () => {
  const committed = (state.emailThreadDraft || [])
    .map((entry) => normaliseThreadMessage(entry))
    .filter(Boolean);
  state.emailThreadDraft = committed;
  return committed;
};

const setEmailThreadDraft = (entries = []) => {
  const normalised = (Array.isArray(entries) ? entries : [])
    .map((entry) => normaliseThreadMessage(entry))
    .filter(Boolean);
  state.emailThreadDraft = normalised;
};

const clampThreadHeight = (value) => {
  const numeric =
    typeof value === "number" ? value : Number.parseFloat(typeof value === "string" ? value : NaN);
  if (!Number.isFinite(numeric)) return null;
  return Math.min(THREAD_MESSAGE_MAX_HEIGHT, Math.max(THREAD_MESSAGE_MIN_HEIGHT, numeric));
};

const autosizeThreadTextarea = (textarea, preferredHeight = null) => {
  if (!textarea) return THREAD_MESSAGE_MIN_HEIGHT;
  let nextHeight;
  if (Number.isFinite(preferredHeight)) {
    nextHeight = clampThreadHeight(preferredHeight) ?? THREAD_MESSAGE_MIN_HEIGHT;
  } else {
    textarea.style.height = "auto";
    nextHeight = clampThreadHeight(textarea.scrollHeight) ?? THREAD_MESSAGE_MIN_HEIGHT;
  }
  textarea.style.height = `${nextHeight}px`;
  return nextHeight;
};

const renderEmailThreadMessages = () => {
  const list = elements.emailThreadList;
  if (!list) return;
  list.replaceChildren();
  const messages = state.emailThreadDraft || [];
  if (!messages.length) {
    const empty = document.createElement("p");
    empty.className = "action-items-empty";
    empty.textContent = "No thread messages captured yet.";
    list.append(empty);
    return;
  }
  messages.forEach((message) => {
    const row = document.createElement("div");
    row.className = "thread-message-row";
    row.dataset.threadMessageId = message.id;
    if (message.completed) {
      row.classList.add("completed");
    }

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.className = "thread-message-checkbox";
    checkbox.dataset.threadMessageId = message.id;
    checkbox.checked = message.completed;

    const textarea = document.createElement("textarea");
    textarea.rows = 2;
    textarea.className = "thread-message-input";
    textarea.dataset.threadMessageId = message.id;
    textarea.value = message.body;
    textarea.placeholder = "Email message details";
    const savedHeight = clampThreadHeight(message.height);
    autosizeThreadTextarea(textarea, savedHeight);

    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className = "ghost-button small";
    removeButton.dataset.action = "email-remove-thread";
    removeButton.dataset.threadMessageId = message.id;
    removeButton.textContent = "Remove";

    row.append(checkbox, textarea, removeButton);
    list.append(row);
  });
};

const addEmailThreadMessages = (entries = []) => {
  const additions = (Array.isArray(entries) ? entries : [entries])
    .map((entry) => {
      if (typeof entry === "string") {
        const value = entry.trim();
        if (!value) return null;
        return { id: generateId("thread-message"), body: value, completed: false, height: null };
      }
      const body = typeof entry?.body === "string" ? entry.body.trim() : "";
      if (!body) return null;
      const height = clampThreadHeight(
        typeof entry.height === "number"
          ? entry.height
          : typeof entry.height === "string"
            ? Number.parseFloat(entry.height)
            : NaN,
      );
      return {
        id: generateId("thread-message"),
        body,
        completed: Boolean(entry.completed),
        height: height ?? null,
      };
    })
    .filter(Boolean);
  if (!additions.length) return;
  state.emailThreadDraft = [...(state.emailThreadDraft || []), ...additions];
  renderEmailThreadMessages();
};

const updateEmailThreadMessage = (id, updates = {}) => {
  if (!id) return;
  state.emailThreadDraft = (state.emailThreadDraft || []).map((message) => {
    if (message.id !== id) return message;
    const updated = normaliseThreadMessage({ ...message, ...updates });
    return updated ?? message;
  });
};

const removeEmailThreadMessage = (id) => {
  if (!id) return;
  state.emailThreadDraft = (state.emailThreadDraft || []).filter((message) => message.id !== id);
  renderEmailThreadMessages();
};

const toggleEmailThreadMessage = (id, completed) => {
  updateEmailThreadMessage(id, { completed });
  renderEmailThreadMessages();
};

const addEmailThreadFromInput = () => {
  const input = elements.emailThreadInput;
  if (!input) return;
  const value = input.value.trim();
  if (!value) return;
  addEmailThreadMessages([value]);
  input.value = "";
};

const handleEmailThreadListChange = (event) => {
  const checkbox = event.target.closest(".thread-message-checkbox");
  if (!checkbox) return;
  const id = checkbox.dataset.threadMessageId;
  toggleEmailThreadMessage(id, checkbox.checked);
};

const handleEmailThreadListInput = (event) => {
  const textarea = event.target.closest(".thread-message-input");
  if (!textarea) return;
  const id = textarea.dataset.threadMessageId;
  const nextHeight = autosizeThreadTextarea(textarea);
  updateEmailThreadMessage(id, { body: textarea.value, height: nextHeight });
};

const handleEmailThreadInputKeydown = (event) => {
  if (event.key !== "Enter" || event.shiftKey) return;
  const input = event.target.closest("[data-email-thread-input]");
  if (!input) return;
  event.preventDefault();
  addEmailThreadFromInput();
};

const handleEmailThreadListPointerUp = (event) => {
  const textarea = event.target.closest(".thread-message-input");
  if (!textarea) return;
  const id = textarea.dataset.threadMessageId;
  const computed = clampThreadHeight(
    Number.parseFloat(window.getComputedStyle(textarea).height || ""),
  );
  if (!id || computed === null) return;
  updateEmailThreadMessage(id, { height: computed });
};

const handleMeetingActionListChange = (event) => {
  const checkbox = event.target.closest(".action-item-checkbox");
  if (!checkbox) return;
  const id = checkbox.dataset.actionItemId;
  updateMeetingActionItem(id, { completed: checkbox.checked });
};

const handleMeetingActionListInput = (event) => {
  const input = event.target.closest(".action-item-input");
  if (!input) return;
  const id = input.dataset.actionItemId;
  state.meetingActionDraft = (state.meetingActionDraft || []).map((item) => {
    if (item.id !== id) return item;
    return { ...item, title: input.value };
  });
};

const handleMeetingActionInputKeydown = (event) => {
  if (event.key !== "Enter") return;
  const input = event.target.closest("[data-meeting-action-input]");
  if (!input) return;
  event.preventDefault();
  const value = input.value.trim();
  if (!value) return;
  addMeetingActionItems(
    value
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
  );
  input.value = "";
};

const addTask = (payload) => {
  const projectId = payload.projectId || "inbox";
  const project = getProjectById(projectId);
  ensureSectionForProject(projectId);
  const sectionId = payload.sectionId || getDefaultSectionId(projectId);
  const now = new Date().toISOString();
  const createdAt = payload.createdAt ?? now;

  const task = {
    id: generateId("task"),
    kind: normaliseTaskKind(payload.kind),
    source: normaliseTaskSource(payload.source),
    title: payload.title,
    description: payload.description,
    dueDate: payload.dueDate,
    priority: payload.priority,
    projectId,
    sectionId,
    companyId: project?.companyId ?? DEFAULT_COMPANY.id,
    departmentId: payload.departmentId || "",
    assigneeId: payload.assigneeId || "",
    attachments: Array.isArray(payload.attachments) ? payload.attachments : [],
    metadata:
      payload.metadata && typeof payload.metadata === "object"
        ? { ...payload.metadata }
        : {},
    links: normaliseTaskLinks(payload.links),
    completed: Boolean(payload.completed ?? false),
    completedAt: payload.completed ? payload.completedAt ?? now : null,
    createdAt,
    updatedAt: now,
    deletedAt: null,
  };

  state.tasks.push(task);
  saveTasks();
  render();
  return task;
};
const updateTask = (taskId, updates) => {
  const index = state.tasks.findIndex((task) => task.id === taskId);
  if (index === -1) return null;
  const previous = state.tasks[index];
  const nextProjectId = updates.projectId ?? previous.projectId;
  const project = getProjectById(nextProjectId);
  ensureSectionForProject(nextProjectId);
  const nextSectionId =
    updates.sectionId && getSectionById(updates.sectionId)
      ? updates.sectionId
      : previous.sectionId || getDefaultSectionId(nextProjectId);

  const nextAttachments = Array.isArray(updates.attachments)
    ? updates.attachments
    : previous.attachments || [];

  const hasSourceUpdate = Object.prototype.hasOwnProperty.call(updates, "source");
  const hasMetadataUpdate = Object.prototype.hasOwnProperty.call(updates, "metadata");
  const hasLinksUpdate = Object.prototype.hasOwnProperty.call(updates, "links");

  const nextKind = normaliseTaskKind(updates.kind ?? previous.kind);
  const nextSource = hasSourceUpdate
    ? normaliseTaskSource(updates.source)
    : normaliseTaskSource(previous.source);
  const nextMetadata = hasMetadataUpdate
    ? updates.metadata && typeof updates.metadata === "object"
      ? { ...updates.metadata }
      : {}
    : previous.metadata && typeof previous.metadata === "object"
    ? { ...previous.metadata }
    : {};
  const nextLinks = hasLinksUpdate
    ? normaliseTaskLinks(updates.links)
    : normaliseTaskLinks(previous.links);

  const completedFlag =
    updates.completed !== undefined ? updates.completed : previous.completed;
  let completedAt = previous.completedAt || null;
  if (updates.completed === true && !previous.completed) {
    completedAt = new Date().toISOString();
  }
  if (updates.completed === false) {
    completedAt = null;
  }

  const updatedTask = {
    ...previous,
    ...updates,
    kind: nextKind,
    source: nextSource,
    metadata: nextMetadata,
    links: nextLinks,
    projectId: nextProjectId,
    sectionId: nextSectionId,
    companyId: project?.companyId ?? previous.companyId ?? DEFAULT_COMPANY.id,
    attachments: nextAttachments,
    completed: completedFlag,
    completedAt,
    updatedAt: new Date().toISOString(),
  };

  state.tasks[index] = updatedTask;
  saveTasks();
  render();
  return updatedTask;
};
const removeTask = (taskId) => {
  const index = state.tasks.findIndex((task) => task.id === taskId);
  if (index === -1) return;
  const previous = state.tasks[index];
  if (previous.deletedAt) return;
  const now = new Date().toISOString();
  state.tasks[index] = { ...previous, deletedAt: now, updatedAt: now };
  saveTasks();
  render();
};

const createProject = (name, companyId = state.activeCompanyId || DEFAULT_COMPANY.id) => {
  const trimmed = name.trim();
  if (!trimmed) return;
  const targetCompany = getCompanyById(companyId) ? companyId : DEFAULT_COMPANY.id;
  const exists = state.projects.some(
    (project) =>
      project.companyId === targetCompany && project.name.toLowerCase() === trimmed.toLowerCase()
  );
  if (exists) {
    window.alert("A project with this name already exists for this company.");
    return;
  }

  const project = {
    id: generateId("project"),
    name: trimmed,
    companyId: targetCompany,
    color: pickProjectColor(state.projects.length),
    createdAt: new Date().toISOString(),
  };

  state.projects.push(project);
  saveProjects();
  ensureSectionForProject(project.id);
  saveSections();
  rememberProjectSelection(project.id);
  setActiveView("project", project.id);
  return project;
};

const createSection = (projectId, name) => {
  const trimmed = name.trim();
  if (!trimmed) return null;
  const sections = getSectionsForProject(projectId);
  const section = {
    id: generateId("section"),
    name: trimmed,
    projectId,
    order: sections.length,
    createdAt: new Date().toISOString(),
  };
  state.sections.push(section);
  saveSections();
  return section;
};

const renameSection = (sectionId, nextName) => {
  const section = getSectionById(sectionId);
  if (!section) return;
  const trimmed = nextName.trim();
  if (!trimmed || trimmed === section.name) return;
  const duplicate = getSectionsForProject(section.projectId).some(
    (entry) => entry.id !== sectionId && entry.name.toLowerCase() === trimmed.toLowerCase()
  );
  if (duplicate) {
    window.alert("Another section in this project already uses that name.");
    return;
  }
  section.name = trimmed;
  section.updatedAt = new Date().toISOString();
  saveSections();
  render();
};

const deleteSection = (sectionId) => {
  const section = getSectionById(sectionId);
  if (!section) return null;

  const sections = getSectionsForProject(section.projectId);
  if (sections.length <= 1) {
    window.alert("A project must have at least one section.");
    return null;
  }

  const fallbackSection = sections.find((entry) => entry.id !== sectionId)?.id;
  state.tasks = state.tasks.map((task) =>
    task.sectionId === sectionId ? { ...task, sectionId: fallbackSection } : task
  );
  state.sections = state.sections.filter((entry) => entry.id !== sectionId);
  saveSections();
  saveTasks();
  render();
  return fallbackSection;
};

const renameProject = (projectId, nextName) => {
  const project = getProjectById(projectId);
  if (!project) return { success: false, error: "Project not found." };
  const trimmed = nextName.trim();
  if (!trimmed) return { success: false, error: "Project name is required." };
  const exists = state.projects.some(
    (entry) =>
      entry.id !== projectId &&
      entry.companyId === project.companyId &&
      entry.name.toLowerCase() === trimmed.toLowerCase(),
  );
  if (exists) {
    return { success: false, error: "Another project in this company already uses that name." };
  }
  project.name = trimmed;
  project.updatedAt = new Date().toISOString();
  saveProjects();
  renderProjectDropdown();
  renderHeader();
  renderTasks();
  return { success: true };
};

const moveProjectToCompany = (projectId, targetCompanyId) => {
  const project = getProjectById(projectId);
  if (!project) return { success: false, error: "Project not found." };
  const company = getCompanyById(targetCompanyId) ?? getCompanyById(DEFAULT_COMPANY.id);
  if (!company) return { success: false, error: "Target company not found." };
  if (project.companyId === company.id) {
    return { success: true, changed: false };
  }

  const now = new Date().toISOString();
  const previousCompanyId = project.companyId;
  project.companyId = company.id;
  project.updatedAt = now;
  state.tasks = state.tasks.map((task) => {
    if (task.projectId !== projectId) return task;
    return { ...task, companyId: company.id, updatedAt: now };
  });

  Object.keys(state.companyRecents).forEach((companyId) => {
    if (state.companyRecents[companyId] === projectId && companyId !== company.id) {
      delete state.companyRecents[companyId];
    }
  });
  state.companyRecents[company.id] = projectId;

  saveProjects();
  saveTasks();
  ensureCompanyPreferences();
  renderProjectDropdown();
  renderCompanyTabs();
  renderHeader();
  renderTasks();

  return { success: true, changed: previousCompanyId !== company.id };
};

const deleteProject = (projectId) => {
  const project = getProjectById(projectId);
  if (!project) return false;
  if (project.isDefault) {
    window.alert("Inbox cannot be deleted.");
    return false;
  }
  const confirmed = window.confirm(`Delete the project "${project.name}" and all of its tasks?`);
  if (!confirmed) return false;

  ensureSectionForProject(DEFAULT_PROJECT.id);
  const defaultSectionId = getDefaultSectionId(DEFAULT_PROJECT.id);
  const company = getCompanyById(project.companyId);
  const now = new Date().toISOString();

  state.tasks = state.tasks.map((task) => {
    if (task.projectId !== projectId) return task;
    const metadata = task.metadata && typeof task.metadata === "object" ? { ...task.metadata } : {};
    metadata.archivedFrom = {
      projectId: project.id,
      projectName: project.name,
      companyId: project.companyId,
      companyName: company?.name ?? "",
      archivedAt: now,
    };
    return {
      ...task,
      projectId: DEFAULT_PROJECT.id,
      sectionId: defaultSectionId,
      companyId: DEFAULT_COMPANY.id,
      metadata,
      updatedAt: now,
    };
  });

  state.sections = state.sections.filter((section) => section.projectId !== projectId);
  state.projects = state.projects.filter((entry) => entry.id !== projectId);
  Object.keys(state.companyRecents).forEach((companyId) => {
    if (state.companyRecents[companyId] === projectId) {
      delete state.companyRecents[companyId];
    }
  });

  saveTasks();
  saveSections();
  saveProjects();

  const remaining = getProjectsForCompany(project.companyId);
  if (state.activeView.type === "project" && state.activeView.value === projectId) {
    if (remaining[0]) {
      setActiveView("project", remaining[0].id);
    } else {
      setActiveView("view", "workspace");
    }
  } else {
    savePreferences();
    render();
  }
  setActiveCompany(ALL_COMPANY_ID);
  return true;
};

const openProjectDialog = (projectId) => {
  if (!elements.projectDialog || !elements.projectForm) return;
  const project = getProjectById(projectId);
  if (!project) return;
  state.editingProjectId = projectId;

  if (elements.projectNameInput) {
    elements.projectNameInput.value = project.name;
  }
  populateProjectCompanyOptions(elements.projectCompanySelect, project.companyId);
  if (elements.projectError) {
    elements.projectError.textContent = "";
  }
  const deleteButton = elements.projectDialog.querySelector('[data-action="delete-project"]');
  if (deleteButton) {
    deleteButton.hidden = Boolean(project.isDefault);
    deleteButton.disabled = Boolean(project.isDefault);
  }

  if (typeof elements.projectDialog.showModal === "function") {
    elements.projectDialog.showModal();
  } else {
    elements.projectDialog.setAttribute("open", "true");
  }
};

const closeProjectDialog = () => {
  state.editingProjectId = null;
  if (elements.projectForm) {
    elements.projectForm.reset();
  }
  if (elements.projectError) {
    elements.projectError.textContent = "";
  }
  if (elements.projectDialog) {
    if (typeof elements.projectDialog.close === "function") {
      elements.projectDialog.close();
    } else {
      elements.projectDialog.removeAttribute("open");
    }
  }
};

const handleProjectFormSubmit = (event) => {
  event.preventDefault();
  if (!state.editingProjectId) return;
  const project = getProjectById(state.editingProjectId);
  if (!project) {
    if (elements.projectError) elements.projectError.textContent = "Project not found.";
    return;
  }

  const nextName = elements.projectNameInput?.value ?? "";
  const targetCompanyId = elements.projectCompanySelect?.value || DEFAULT_COMPANY.id;
  const errors = [];
  let changed = false;

  if (targetCompanyId !== project.companyId) {
    const moveResult = moveProjectToCompany(project.id, targetCompanyId);
    if (!moveResult.success) {
      errors.push(moveResult.error);
    } else {
      changed = moveResult.changed || changed;
    }
  }

  if (nextName.trim() && nextName.trim() !== project.name) {
    const renameResult = renameProject(project.id, nextName);
    if (!renameResult.success) {
      errors.push(renameResult.error);
    } else {
      changed = true;
    }
  }

  if (!nextName.trim()) {
    errors.push("Project name is required.");
  }

  if (errors.length) {
    if (elements.projectError) elements.projectError.textContent = errors.join(" ");
    return;
  }

  if (!changed) {
    if (elements.projectError) elements.projectError.textContent = "No changes to save.";
    return;
  }

  const updated = getProjectById(project.id);
  if (updated) {
    setActiveView("project", updated.id);
  }
  closeProjectDialog();
};

const handleProjectFormClick = (event) => {
  const action = event.target.dataset.action;
  if (action === "close-project") {
    closeProjectDialog();
    event.preventDefault();
    return;
  }
  if (action === "delete-project" && state.editingProjectId) {
    event.preventDefault();
    const project = getProjectById(state.editingProjectId);
    if (!project) return;
    const confirmed = window.confirm(
      `Delete the project "${project.name}"? All tasks will be moved to the workspace inbox.`,
    );
    if (confirmed) {
      deleteProject(project.id);
      closeProjectDialog();
    }
  }
};


const createCompany = (name) => {
  const trimmed = name.trim();
  if (!trimmed) return null;
  const exists = state.companies.some(
    (company) => company.name.toLowerCase() === trimmed.toLowerCase(),
  );
  if (exists) {
    window.alert("A company with this name already exists.");
    return null;
  }
  const company = {
    id: generateId("company"),
    name: trimmed,
    createdAt: new Date().toISOString(),
  };
  state.companies.push(company);
  state.activeCompanyId = company.id;
  state.companyRecents[company.id] = "";
  saveCompanies();
  savePreferences();
  renderCompanyTabs();
  renderProjectDropdown();
  return company;
};

const deleteCompany = (companyId) => {
  const company = getCompanyById(companyId);
  if (!company) return false;
  if (company.isDefault) {
    window.alert("The default company cannot be deleted.");
    return false;
  }
  const confirmed = window.confirm(
    `Delete the company "${company.name}"? All related projects, sections, and tasks will be removed.`,
  );
  if (!confirmed) return false;

  const projectIds = state.projects
    .filter((project) => project.companyId === companyId)
    .map((project) => project.id);

  ensureSectionForProject(DEFAULT_PROJECT.id);
  const defaultSectionId = getDefaultSectionId(DEFAULT_PROJECT.id);
  const now = new Date().toISOString();
  const companyLookup = getCompanyById(companyId);
  const projectLookup = new Map(state.projects.map((project) => [project.id, project]));

  state.tasks = state.tasks.map((task) => {
    if (!projectIds.includes(task.projectId)) return task;
    const originProject = projectLookup.get(task.projectId);
    const metadata = task.metadata && typeof task.metadata === "object" ? { ...task.metadata } : {};
    metadata.archivedFrom = {
      projectId: originProject?.id ?? "",
      projectName: originProject?.name ?? "",
      companyId,
      companyName: companyLookup?.name ?? "",
      archivedAt: now,
    };
    return {
      ...task,
      projectId: DEFAULT_PROJECT.id,
      sectionId: defaultSectionId,
      companyId: DEFAULT_COMPANY.id,
      metadata,
      updatedAt: now,
    };
  });

  state.sections = state.sections.filter((section) => !projectIds.includes(section.projectId));
  state.projects = state.projects.filter((project) => project.companyId !== companyId);
  state.companies = state.companies.filter((entry) => entry.id !== companyId);
  delete state.companyRecents[companyId];

  saveTasks();
  saveSections();
  saveProjects();
  saveCompanies();

  if (state.activeCompanyId === companyId) {
    state.activeCompanyId = state.companies[0]?.id ?? DEFAULT_COMPANY.id;
    const fallbackCompany = getCompanyById(state.activeCompanyId);
    if (fallbackCompany) {
      const fallbackProject = getProjectsForCompany(fallbackCompany.id)[0];
      if (fallbackProject) {
        setActiveView("project", fallbackProject.id);
      } else {
        savePreferences();
        render();
      }
    } else {
      setActiveView("view", "workspace");
    }
  } else {
    savePreferences();
    render();
  }
  setActiveCompany(ALL_COMPANY_ID);
  return true;
};

const addMember = (name, departmentId) => {
  const trimmed = name.trim();
  if (!trimmed) return;
  const member = {
    id: generateId("member"),
    name: trimmed,
    departmentId: departmentId || "",
    title: "",
    email: "",
    avatarUrl: "",
    createdAt: new Date().toISOString(),
  };
  state.members.push(member);
  saveMembers();
  updateTeamSelects();
  renderMemberList();
  renderTeamStatus();
  updateDashboardMetrics();
  return member;
};

const removeMember = (memberId) => {
  state.members = state.members.filter((member) => member.id !== memberId);
  state.tasks = state.tasks.map((task) =>
    task.assigneeId === memberId ? { ...task, assigneeId: "" } : task
  );
  saveMembers();
  saveTasks();
  updateTeamSelects();
  renderMemberList();
  renderTeamStatus();
  updateDashboardMetrics();
  renderTasks();
};

const addDepartment = (name) => {
  const trimmed = name.trim();
  if (!trimmed) return;
  const exists = state.departments.some(
    (department) => department.name.toLowerCase() === trimmed.toLowerCase()
  );
  if (exists) {
    window.alert("A department with this name already exists.");
    return;
  }
  const department = {
    id: generateId("department"),
    name: trimmed,
    isDefault: false,
    createdAt: new Date().toISOString(),
  };
  state.departments.push(department);
  saveDepartments();
  updateTeamSelects();
  renderDepartmentList();
  renderTeamStatus();
  updateDashboardMetrics();
  return department;
};


const removeDepartment = (departmentId) => {
  const department = getDepartmentById(departmentId);
  if (!department || department.isDefault) {
    window.alert("The default department cannot be removed.");
    return;
  }

  state.departments = state.departments.filter((entry) => entry.id !== departmentId);
  state.members = state.members.map((member) =>
    member.departmentId === departmentId ? { ...member, departmentId: "" } : member
  );
  state.tasks = state.tasks.map((task) =>
    task.departmentId === departmentId ? { ...task, departmentId: "" } : task
  );
  saveDepartments();
  saveMembers();
  saveTasks();
  updateTeamSelects();
  renderDepartmentList();
  renderMemberList();
  renderTeamStatus();
  renderTasks();
};


const renderMemberList = () => {
  if (!elements.memberList) return;
  const fragment = document.createDocumentFragment();
  state.members.forEach((member) => {
    const item = document.createElement("li");
    const label = document.createElement("span");
    label.textContent = formatMemberLabel(member);
    item.append(label);

    const removeBtn = document.createElement("button");
    removeBtn.className = "ghost-button small";
    removeBtn.type = "button";
    removeBtn.dataset.action = "remove-member";
    removeBtn.dataset.memberId = member.id;
    removeBtn.textContent = "Remove";
    item.append(removeBtn);

    fragment.append(item);
  });

  if (!state.members.length) {
    const empty = document.createElement("li");
    empty.textContent = "No members yet.";
    fragment.append(empty);
  }

  elements.memberList.replaceChildren(fragment);
};

const renderDepartmentList = () => {
  if (!elements.departmentList) return;
  const fragment = document.createDocumentFragment();
  state.departments.forEach((department) => {
    const item = document.createElement("li");
    const label = document.createElement("span");
    label.textContent = department.name;
    item.append(label);

    if (!department.isDefault) {
      const removeBtn = document.createElement("button");
      removeBtn.className = "ghost-button small";
      removeBtn.type = "button";
      removeBtn.dataset.action = "remove-department";
      removeBtn.dataset.departmentId = department.id;
      removeBtn.textContent = "Remove";
      item.append(removeBtn);
    }

    fragment.append(item);
  });

  elements.departmentList.replaceChildren(fragment);
};
const openMembersDialog = () => {
  renderMemberList();
  updateTeamSelects();
  if (typeof elements.membersDialog.showModal === "function") {
    elements.membersDialog.showModal();
  } else {
    elements.membersDialog.setAttribute("open", "true");
  }
};

const closeMembersDialog = () => {
  elements.membersForm.reset();
  if (typeof elements.membersDialog.close === "function") {
    elements.membersDialog.close();
  } else {
    elements.membersDialog.removeAttribute("open");
  }
};

const openDepartmentsDialog = () => {
  renderDepartmentList();
  if (typeof elements.departmentsDialog.showModal === "function") {
    elements.departmentsDialog.showModal();
  } else {
    elements.departmentsDialog.setAttribute("open", "true");
  }
};

const closeDepartmentsDialog = () => {
  elements.departmentsForm.reset();
  if (typeof elements.departmentsDialog.close === "function") {
    elements.departmentsDialog.close();
  } else {
    elements.departmentsDialog.removeAttribute("open");
  }
};

const handleViewClick = (event) => {
  const button = event.target.closest("[data-view]");
  if (!button) return;
  setActiveView("view", button.dataset.view);
};

const openTaskEditor = (task) => {
  if (!task) return;
  const kind = task.kind ?? "task";
  if (kind === "meeting") {
    openMeetingDialog(task.projectId, task);
    return;
  }
  if (kind === "email") {
    openEmailDialog(task.projectId, task);
    return;
  }
  openTaskDialog(task.id);
};

const handleTaskActionClick = (event) => {
  const button = event.target.closest("[data-action]");
  if (!button) return;
  const action = button.dataset.action;
  if (action === "project-overview-expand") {
    expandProjectOverviewSection(button.dataset.projectId, button.dataset.groupId, button.dataset.total);
    return;
  }
  if (action === "edit-section") {
    const sectionId = button.dataset.sectionId;
    if (sectionId) {
      openSectionEditor(sectionId);
    }
    return;
  }
  if (action === "edit") {
    const item = button.closest(".task-item");
    if (!item) return;
    const task = state.tasks.find((entry) => entry.id === item.dataset.taskId);
    if (!task) return;
    openTaskEditor(task);
    return;
  }
  if (action === "delete") {
    const item = button.closest(".task-item");
    if (!item) return;
    const taskId = item.dataset.taskId;
    const confirmed = window.confirm("Delete this task? This cannot be undone.");
    if (confirmed) {
      removeTask(taskId);
    }
  }
};

const handleTaskItemOpen = (event) => {
  if (state.viewMode !== "list") return;
  const item = event.target.closest(".task-item");
  if (!item) return;
  if (
    event.target.closest(".task-actions") ||
    event.target.closest(".task-checkbox") ||
    event.target.closest("[data-action]") ||
    event.target.closest("button") ||
    event.target.closest("input") ||
    event.target.closest("a")
  ) {
    return;
  }
  const task = state.tasks.find((entry) => entry.id === item.dataset.taskId);
  if (!task || task.deletedAt) return;
  openTaskEditor(task);
};

const prepareMeetingDialog = (projectId, task = null) => {
  const form = elements.meetingForm;
  if (!form) return;
  form.reset();
  const project = getProjectById(projectId);
  if (elements.meetingProjectLabel) {
    elements.meetingProjectLabel.textContent = project ? project.name : "Inbox";
  }
  form.elements.projectId.value = projectId || "inbox";
  if (elements.meetingDepartment) {
    populateDepartmentOptions(elements.meetingDepartment, task?.departmentId || "");
  }
  if (elements.meetingAssignee) {
    populateMemberOptions(
      elements.meetingAssignee,
      task?.assigneeId ?? "",
      elements.meetingDepartment?.value ?? ""
    );
  }
  if (elements.meetingPriority) {
    elements.meetingPriority.value = task?.priority ?? "medium";
  }
  form.elements.date.value = task?.dueDate ?? "";
  form.elements.meetingType.value = task?.metadata?.meetingType ?? "";
  form.elements.attendees.value = task?.metadata?.attendees ?? "";
  form.elements.title.value = task?.title ?? "";
  if (form.elements.notes) {
    form.elements.notes.value = task?.description ?? "";
  }

  const actionItems = Array.isArray(task?.metadata?.actionItems) ? task.metadata.actionItems : [];
  setMeetingActionDraft(actionItems);
  renderMeetingActionItems();
  if (elements.meetingActionInput) {
    elements.meetingActionInput.value = "";
  }
  resetLinkList(elements.meetingLinksList, Array.isArray(task?.links) ? task.links : []);

  state.meetingCompletedDraft = Boolean(task?.completed);
  const completedAt = task?.completedAt ?? null;
  const isEditing = Boolean(task);
  if (elements.meetingDialogToggle) {
    elements.meetingDialogToggle.hidden = !isEditing;
    elements.meetingDialogToggle.disabled = !isEditing;
  }
  updateMeetingDialogCompletionState(state.meetingCompletedDraft, completedAt);
  if (!isEditing && elements.meetingDialogStatus) {
    elements.meetingDialogStatus.textContent = "";
  }

  if (elements.meetingError) elements.meetingError.textContent = "";
  registerAutosizeTextarea(form.elements.attendees);
  registerAutosizeTextarea(form.elements.notes);
  form.elements.attendees?.dispatchEvent(new Event("input", { bubbles: false }));
  form.elements.notes?.dispatchEvent(new Event("input", { bubbles: false }));
};

const openMeetingDialog = (projectId, task = null) => {
  const dialog = elements.meetingDialog;
  if (!dialog || !elements.meetingForm) return;
  prepareMeetingDialog(projectId, task);
  state.editingMeetingId = task?.id ?? null;
  if (typeof dialog.showModal === "function") dialog.showModal();
  else dialog.setAttribute("open", "true");
  window.requestAnimationFrame(() => {
    elements.meetingForm?.elements.title?.focus();
  });
};

const closeMeetingDialog = () => {
  const dialog = elements.meetingDialog;
  const form = elements.meetingForm;
  state.editingMeetingId = null;
  state.meetingCompletedDraft = false;
  if (form) {
    form.reset();
    resetLinkList(elements.meetingLinksList);
    setMeetingActionDraft([]);
    renderMeetingActionItems();
    if (elements.meetingError) elements.meetingError.textContent = "";
    applySavedTextareaHeight(form.elements.attendees);
    applySavedTextareaHeight(form.elements.notes);
  }
  if (elements.meetingActionInput) {
    elements.meetingActionInput.value = "";
  }
  if (elements.meetingDialogToggle) {
    elements.meetingDialogToggle.hidden = true;
    elements.meetingDialogToggle.disabled = true;
    elements.meetingDialogToggle.dataset.completedState = "false";
    elements.meetingDialogToggle.textContent = "Mark completed";
  }
  if (elements.meetingDialogStatus) {
    elements.meetingDialogStatus.textContent = "";
  }
  if (dialog) {
    if (typeof dialog.close === "function") dialog.close();
    else dialog.removeAttribute("open");
  }
};

const commitMeetingForm = ({ allowCreate = false } = {}) => {
  const form = elements.meetingForm;
  if (!form) return false;
  const isEditing = Boolean(state.editingMeetingId);
  if (!isEditing && !allowCreate) {
    if (elements.meetingError) elements.meetingError.textContent = "";
    return true;
  }

  const titleField = form.elements.title;
  const title = titleField?.value?.trim() ?? "";
  if (!title) {
    if (elements.meetingError) {
      elements.meetingError.textContent = "Title is required.";
    }
    titleField?.focus();
    return false;
  }

  const projectId = form.elements.projectId.value || "inbox";
  ensureSectionForProject(projectId);
  const sectionId = getDefaultSectionId(projectId);
  const links = collectLinks(elements.meetingLinksList);
  const existing = state.tasks.find((task) => task.id === state.editingMeetingId);
  const metadata =
    existing?.metadata && typeof existing.metadata === "object" ? { ...existing.metadata } : {};
  metadata.meetingType = form.elements.meetingType.value;
  metadata.attendees = form.elements.attendees.value.trim();
  const actionItems = commitMeetingActionDraft();
  metadata.actionItems = actionItems;
  metadata.links = links;

  const payload = {
    title,
    description: form.elements.notes?.value?.trim() ?? "",
    dueDate: form.elements.date.value || "",
    priority: form.elements.priority.value || "medium",
    projectId,
    sectionId,
    departmentId: form.elements.department.value || "",
    assigneeId: form.elements.assignee?.value || "",
    kind: "meeting",
    source: existing?.source ?? "manual",
    metadata,
    links,
    actionItems,
    completed: state.meetingCompletedDraft,
  };

  if (isEditing && existing) {
    const updated = updateTask(state.editingMeetingId, payload);
    if (!updated) return false;
    state.meetingCompletedDraft = updated.completed;
    updateMeetingDialogCompletionState(updated.completed, updated.completedAt ?? null);
  } else if (allowCreate) {
    addTask(payload);
  }

  if (elements.meetingError) elements.meetingError.textContent = "";
  return true;
};

const autoSaveMeetingDialog = () => {
  const saved = commitMeetingForm({ allowCreate: false });
  if (saved) {
    closeMeetingDialog();
  }
  return saved;
};

const handleMeetingFormSubmit = (event) => {
  event.preventDefault();
  commitMeetingForm({ allowCreate: true }) && closeMeetingDialog();
};

const handleMeetingFormClick = (event) => {
  const action = event.target.dataset.action;
  if (action === "meeting-toggle-completion") {
    event.preventDefault();
    const nextCompleted = !state.meetingCompletedDraft;
    state.meetingCompletedDraft = nextCompleted;
    if (state.editingMeetingId) {
      const updated = updateTask(state.editingMeetingId, { completed: nextCompleted });
      if (updated) {
        state.meetingCompletedDraft = updated.completed;
        updateMeetingDialogCompletionState(updated.completed, updated.completedAt ?? null);
      }
    } else {
      updateMeetingDialogCompletionState(
        nextCompleted,
        nextCompleted ? new Date().toISOString() : null,
      );
    }
    return;
  }
  if (action === "meeting-add-link") {
    event.preventDefault();
    createLinkRow(elements.meetingLinksList);
    return;
  }
  if (action === "remove-link") {
    event.preventDefault();
    const row = event.target.closest('[data-link-row]');
    row?.remove();
    if (elements.meetingLinksList && elements.meetingLinksList.childElementCount === 0) {
      createLinkRow(elements.meetingLinksList);
    }
    return;
  }
  if (action === "meeting-add-action") {
    event.preventDefault();
    const input = elements.meetingActionInput;
    if (!input) return;
    const value = input.value.trim();
    if (!value) return;
    addMeetingActionItems(
      value
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter(Boolean),
    );
    input.value = "";
    input.focus();
    return;
  }
  if (action === "meeting-remove-action") {
    event.preventDefault();
    const id = event.target.dataset.actionItemId;
    removeMeetingActionItem(id);
    return;
  }
  if (action === "close-meeting") {
    event.preventDefault();
    autoSaveMeetingDialog();
  }
};

const prepareEmailDialog = (projectId, task = null) => {
  const form = elements.emailForm;
  if (!form) return;
  form.reset();
  const project = getProjectById(projectId);
  if (elements.emailProjectLabel) {
    elements.emailProjectLabel.textContent = project ? project.name : "Inbox";
  }
  form.elements.projectId.value = projectId || "inbox";
  if (elements.emailDepartment) {
    populateDepartmentOptions(elements.emailDepartment, task?.departmentId || "");
  }
  if (elements.emailAssignee) {
    populateMemberOptions(
      elements.emailAssignee,
      task?.assigneeId ?? "",
      elements.emailDepartment?.value ?? ""
    );
  }
  if (elements.emailPriority) {
    elements.emailPriority.value = task?.priority ?? "medium";
  }
  if (elements.emailStatus) {
    elements.emailStatus.value = task?.metadata?.status ?? "Pending";
  }
  form.elements.date.value = task?.dueDate ?? "";
  form.elements.emailAddress.value = task?.metadata?.emailAddress ?? "";
  form.elements.title.value = task?.title ?? "";
  if (form.elements.notes) {
    form.elements.notes.value = task?.description ?? "";
  }
  resetLinkList(elements.emailLinksList, Array.isArray(task?.links) ? task.links : []);
  const threadMessages = Array.isArray(task?.metadata?.threadMessages)
    ? task.metadata.threadMessages
    : [];
  setEmailThreadDraft(threadMessages);
  renderEmailThreadMessages();
  if (elements.emailThreadInput) {
    elements.emailThreadInput.value = "";
  }
  state.emailCompletedDraft = Boolean(task?.completed);
  const completedAt = task?.completedAt ?? null;
  const isEditing = Boolean(task);
  if (elements.emailDialogToggle) {
    elements.emailDialogToggle.hidden = !isEditing;
    elements.emailDialogToggle.disabled = !isEditing;
  }
  updateEmailDialogCompletionState(state.emailCompletedDraft, completedAt);
  if (!isEditing && elements.emailDialogStatus) {
    elements.emailDialogStatus.textContent = "";
  }
  if (elements.emailError) elements.emailError.textContent = "";
  registerAutosizeTextarea(form.elements.notes);
  form.elements.notes?.dispatchEvent(new Event("input", { bubbles: false }));
};

const openEmailDialog = (projectId, task = null) => {
  const dialog = elements.emailDialog;
  if (!dialog || !elements.emailForm) return;
  prepareEmailDialog(projectId, task);
  state.editingEmailId = task?.id ?? null;
  if (typeof dialog.showModal === "function") dialog.showModal();
  else dialog.setAttribute("open", "true");
  window.requestAnimationFrame(() => {
    elements.emailForm?.elements.title?.focus();
  });
};

const closeEmailDialog = () => {
  const dialog = elements.emailDialog;
  const form = elements.emailForm;
  state.editingEmailId = null;
  state.emailCompletedDraft = false;
  if (form) {
    form.reset();
    resetLinkList(elements.emailLinksList);
    setEmailThreadDraft([]);
    renderEmailThreadMessages();
    if (elements.emailThreadInput) {
      elements.emailThreadInput.value = "";
    }
    if (elements.emailError) elements.emailError.textContent = "";
    applySavedTextareaHeight(form.elements.notes);
  }
  if (elements.emailDialogToggle) {
    elements.emailDialogToggle.hidden = true;
    elements.emailDialogToggle.disabled = true;
    elements.emailDialogToggle.dataset.completedState = "false";
    elements.emailDialogToggle.textContent = "Mark completed";
  }
  if (elements.emailDialogStatus) {
    elements.emailDialogStatus.textContent = "";
  }
  if (dialog) {
    if (typeof dialog.close === "function") dialog.close();
    else dialog.removeAttribute("open");
  }
};

const commitEmailForm = ({ allowCreate = false } = {}) => {
  const form = elements.emailForm;
  if (!form) return false;
  const isEditing = Boolean(state.editingEmailId);
  if (!isEditing && !allowCreate) {
    if (elements.emailError) elements.emailError.textContent = "";
    return true;
  }

  const titleField = form.elements.title;
  const title = titleField?.value?.trim() ?? "";
  if (!title) {
    if (elements.emailError) {
      elements.emailError.textContent = "Title is required.";
    }
    titleField?.focus();
    return false;
  }

  const projectId = form.elements.projectId.value || "inbox";
  ensureSectionForProject(projectId);
  const sectionId = getDefaultSectionId(projectId);
  const links = collectLinks(elements.emailLinksList);
  const existing = state.tasks.find((task) => task.id === state.editingEmailId);
  const metadata =
    existing?.metadata && typeof existing.metadata === "object" ? { ...existing.metadata } : {};
  metadata.emailAddress = form.elements.emailAddress.value.trim();
  metadata.status = form.elements.status.value;
  metadata.links = links;
  const threadMessages = commitEmailThreadDraft();
  metadata.threadMessages = threadMessages;

  const payload = {
    title,
    description: form.elements.notes?.value?.trim() ?? "",
    dueDate: form.elements.date.value || "",
    priority: form.elements.priority.value || "medium",
    projectId,
    sectionId,
    departmentId: form.elements.department.value || "",
    assigneeId: form.elements.assignee?.value || "",
    kind: "email",
    source: existing?.source ?? "manual",
    metadata,
    links,
    completed: state.emailCompletedDraft,
  };

  if (isEditing && existing) {
    const updated = updateTask(state.editingEmailId, payload);
    if (!updated) return false;
    state.emailCompletedDraft = updated.completed;
    updateEmailDialogCompletionState(updated.completed, updated.completedAt ?? null);
  } else if (allowCreate) {
    addTask(payload);
  }

  if (elements.emailError) elements.emailError.textContent = "";
  return true;
};

const autoSaveEmailDialog = () => {
  const saved = commitEmailForm({ allowCreate: false });
  if (saved) {
    closeEmailDialog();
  }
  return saved;
};

const handleEmailFormSubmit = (event) => {
  event.preventDefault();
  commitEmailForm({ allowCreate: true }) && closeEmailDialog();
};

const handleEmailFormClick = (event) => {
  const action = event.target.dataset.action;
  if (action === "email-toggle-completion") {
    event.preventDefault();
    const nextCompleted = !state.emailCompletedDraft;
    state.emailCompletedDraft = nextCompleted;
    if (state.editingEmailId) {
      const updated = updateTask(state.editingEmailId, { completed: nextCompleted });
      if (updated) {
        state.emailCompletedDraft = updated.completed;
        updateEmailDialogCompletionState(updated.completed, updated.completedAt ?? null);
      }
    } else {
      updateEmailDialogCompletionState(
        nextCompleted,
        nextCompleted ? new Date().toISOString() : null,
      );
    }
    return;
  }
  if (action === "email-add-link") {
    event.preventDefault();
    createLinkRow(elements.emailLinksList);
    return;
  }
  if (action === "email-add-thread") {
    event.preventDefault();
    addEmailThreadFromInput();
    return;
  }
  if (action === "remove-link") {
    event.preventDefault();
    const row = event.target.closest('[data-link-row]');
    row?.remove();
    if (elements.emailLinksList && elements.emailLinksList.childElementCount === 0) {
      createLinkRow(elements.emailLinksList);
    }
    return;
  }
  if (action === "email-remove-thread") {
    event.preventDefault();
    const threadId = event.target.dataset.threadMessageId;
    removeEmailThreadMessage(threadId);
    return;
  }
  if (action === "close-email") {
    event.preventDefault();
    autoSaveEmailDialog();
  }
};

const handleQuickAddSubmit = async (event) => {
  event.preventDefault();
  const data = new FormData(elements.quickAddForm);
  const title = normaliseTitle(data.get("title") ?? "");
  if (!title) {
    elements.quickAddError.textContent = "Please provide a task title.";
    return;
  }
  elements.quickAddError.textContent = "";

  const projectId = data.get("project") || "inbox";
  const sectionId = data.get("section") || getDefaultSectionId(projectId);
  const links = collectLinks(elements.quickAddLinksList);
  const attachments = elements.quickAddAttachments ? await readFilesAsData(elements.quickAddAttachments.files) : [];

  try {
    addTask({
      title,
      description: (data.get("description") ?? "").trim(),
      dueDate: data.get("dueDate") || "",
      priority: data.get("priority") || "medium",
      projectId,
      sectionId,
      departmentId: data.get("department") || "",
      assigneeId: data.get("assignee") || "",
      attachments,
      links,
    });

    resetQuickAddForm();
    closeQuickAddForm();
  } catch (error) {
    console.error("Failed to create task.", error);
    elements.quickAddError.textContent = "Unable to create task. Please try again.";
  }
};

const handleQuickAddClick = (event) => {
  const action = event.target.dataset.action;
  if (action === "quick-add-link") {
    event.preventDefault();
    if (elements.quickAddLinksList) {
      createLinkRow(elements.quickAddLinksList);
    }
    return;
  }
  if (action === "remove-link") {
    const row = event.target.closest('[data-link-row]');
    if (row && elements.quickAddLinksList?.contains(row)) {
      event.preventDefault();
      row.remove();
      if (elements.quickAddLinksList.childElementCount === 0) {
        createLinkRow(elements.quickAddLinksList);
      }
    }
  }
};

const handleQuickAddCancel = () => {
  resetQuickAddForm();
  closeQuickAddForm();
};

const resetQuickAddForm = () => {
  if (!elements.quickAddForm) return;
  elements.quickAddForm.reset();
  if (elements.quickAddPriority) elements.quickAddPriority.value = "medium";
  if (elements.quickAddDepartment) elements.quickAddDepartment.value = "";
  if (elements.quickAddAssignee) elements.quickAddAssignee.value = "";
  elements.quickAddError.textContent = "";
  if (elements.quickAddAttachmentList) elements.quickAddAttachmentList.replaceChildren();
  if (elements.quickAddAttachments) elements.quickAddAttachments.value = "";
  resetLinkList(elements.quickAddLinksList);
  syncQuickAddSelectors();
  const descriptionField = elements.quickAddForm?.elements?.description;
  if (descriptionField) {
    applySavedTextareaHeight(descriptionField);
  }
};

const openQuickAddForm = () => {
  if (!elements.quickAddForm) return;
  elements.quickAddForm.classList.remove("hidden");
  state.isQuickAddOpen = true;
  const titleField = elements.quickAddForm.querySelector('input[name=\"title\"]');
  if (titleField) {
    titleField.focus();
  }
};

const closeQuickAddForm = () => {
  if (!elements.quickAddForm) return;
  elements.quickAddForm.classList.add("hidden");
  state.isQuickAddOpen = false;
  if (elements.toggleQuickAdd) {
    elements.toggleQuickAdd.textContent = "Add Task";
  }
};

const toggleQuickAddForm = () => {
  if (!elements.quickAddForm) return;
  if (state.isQuickAddOpen) {
    closeQuickAddForm();
  } else {
    openQuickAddForm();
  }
};

const updateQuickAddAttachmentPreview = () => {
  if (!elements.quickAddAttachmentList || !elements.quickAddAttachments) return;
  elements.quickAddAttachmentList.replaceChildren();
  const files = [...elements.quickAddAttachments.files];
  if (!files.length) return;
  files.forEach((file) => {
    const chip = document.createElement("li");
    chip.className = "attachment-chip";
    chip.textContent = file.name;
    elements.quickAddAttachmentList.append(chip);
  });
};

const cloneAttachments = (attachments = []) => attachments.map((attachment) => ({ ...attachment }));

const renderDialogAttachments = () => {
  if (!elements.dialogAttachmentList) return;
  elements.dialogAttachmentList.replaceChildren();
  if (!state.dialogAttachmentDraft.length) return;
  state.dialogAttachmentDraft.forEach((attachment) => {
    const item = document.createElement('li');
    item.className = 'attachment-chip';
    const link = document.createElement('a');
    link.href = attachment.data;
    link.textContent = attachment.name;
    link.download = attachment.name;
    link.target = '_blank';
    link.rel = 'noopener';
    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.dataset.action = 'remove-dialog-attachment';
    removeBtn.dataset.attachmentId = attachment.id;
    removeBtn.textContent = 'x';
    item.append(link, removeBtn);
    elements.dialogAttachmentList.append(item);
  });
};

const handleDialogAttachmentsInput = async (event) => {
  const files = await readFilesAsData(event.target.files);
  if (!files.length) return;
  state.dialogAttachmentDraft.push(...files);
  renderDialogAttachments();
  event.target.value = '';
};

const removeDialogAttachment = (attachmentId) => {
  state.dialogAttachmentDraft = state.dialogAttachmentDraft.filter((attachment) => attachment.id !== attachmentId);
  renderDialogAttachments();
};

const focusQuickAdd = () => {
  if (!elements.quickAddForm) return;
  elements.quickAddForm.scrollIntoView({ behavior: "smooth", block: "center" });
  const titleField = elements.quickAddForm.querySelector('input[name="title"]');
  if (titleField) {
    titleField.focus();
  }
};

const handleSearchInput = (event) => {
  setSearchTerm(event.target.value);
};

const handleAddProject = () => {
  if (!state.activeCompanyId) {
    window.alert("Create a company before adding projects.");
    return;
  }
  const name = window.prompt("Project name");
  if (!name) return;
  try {
    createProject(name, state.activeCompanyId);
  } catch (error) {
    console.error("Failed to add project.", error);
  }
};

const handleExport = () => {
  const payload = {
    tasks: state.tasks,
    projects: state.projects,
    sections: state.sections,
    companies: state.companies,
    members: state.members,
    departments: state.departments,
    userguide: state.userguide,
    exportedAt: new Date().toISOString(),
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = "synergy-tasks-backup.json";
  anchor.rel = "noopener";
  anchor.click();
  URL.revokeObjectURL(url);
};

const handleClearAll = () => {
  const confirmed = window.confirm("Delete all tasks, projects, sections, and team data? This cannot be undone.");
  if (!confirmed) return;

  state.tasks = [];
  state.projects = [{ ...DEFAULT_PROJECT }];
  state.sections = [];
  state.companies = [{ ...DEFAULT_COMPANY }];
  state.members = [];
  state.departments = [{ ...DEFAULT_DEPARTMENT }];
  state.activeCompanyId = DEFAULT_COMPANY.id;
  state.companyRecents = {};
  state.userguide = [...DEFAULT_USERGUIDE];
  state.imports = { whatsapp: {} };

  ensureSectionForProject("inbox");

  saveTasks();
  saveProjects();
  saveSections();
  saveCompanies();
  saveMembers();
  saveDepartments();
  saveUserguide();
  saveImports();
  setActiveView("view", "workspace");
};

const openTaskDialog = (taskId) => {
  const task = state.tasks.find((entry) => entry.id === taskId);
  if (!task) return;
  state.editingTaskId = taskId;

  populateProjectOptions();
  updateTeamSelects();
  populateSectionOptions(elements.dialogForm.elements.section, task.projectId, task.sectionId);

  elements.dialogForm.title.value = task.title;
  elements.dialogForm.description.value = task.description ?? "";
  registerAutosizeTextarea(elements.dialogForm.description);
  elements.dialogForm.description?.dispatchEvent(new Event("input", { bubbles: false }));
  elements.dialogForm.dueDate.value = task.dueDate ?? "";
  elements.dialogForm.priority.value = task.priority ?? "medium";
  elements.dialogForm.project.value = task.projectId ?? "inbox";
  elements.dialogForm.section.value = task.sectionId ?? getDefaultSectionId(task.projectId);
  elements.dialogForm.department.value = task.departmentId ?? "";
  populateMemberOptions(
    elements.dialogForm.assignee,
    task.assigneeId ?? "",
    elements.dialogForm.department.value
  );
  elements.dialogForm.assignee.value = task.assigneeId ?? "";
  elements.dialogForm.completed.checked = Boolean(task.completed);
  state.dialogAttachmentDraft = cloneAttachments(task.attachments || []);
  renderDialogAttachments();
  if (elements.dialogAttachmentsInput) {
    elements.dialogAttachmentsInput.value = "";
  }
  resetLinkList(elements.dialogLinksList, Array.isArray(task?.links) ? task.links : []);

  updateTaskDialogCompletionState(task);

  if (typeof elements.taskDialog.showModal === "function") {
    elements.taskDialog.showModal();
  } else {
    elements.taskDialog.setAttribute("open", "true");
  }
};

const closeTaskDialog = () => {
  state.editingTaskId = null;
  state.dialogAttachmentDraft = [];
  if (elements.dialogAttachmentList) {
    elements.dialogAttachmentList.replaceChildren();
  }
  if (elements.dialogAttachmentsInput) {
    elements.dialogAttachmentsInput.value = "";
  }
  resetLinkList(elements.dialogLinksList);
  if (elements.dialogForm?.elements?.description) {
    applySavedTextareaHeight(elements.dialogForm.elements.description);
  }
  if (elements.taskDialogStatus) {
    elements.taskDialogStatus.textContent = "";
  }
  if (elements.taskDialogToggle) {
    elements.taskDialogToggle.dataset.completedState = "";
    elements.taskDialogToggle.textContent = "Mark completed";
  }
  if (typeof elements.taskDialog.close === "function") {
    elements.taskDialog.close();
  } else {
    elements.taskDialog.removeAttribute("open");
  }
};

const focusTaskRow = (taskId) => {
  const row = document.querySelector(`.task-item[data-task-id="${taskId}"]`);
  if (!row) return;
  row.scrollIntoView({ block: "center", behavior: "smooth" });
  row.classList.add("task-item--focus");
  window.setTimeout(() => row.classList.remove("task-item--focus"), 1200);
};

const navigateToTask = (taskId) => {
  const task = state.tasks.find((entry) => entry.id === taskId);
  if (!task) return;
  const project = getProjectById(task.projectId);
  const companyId = project?.companyId ?? DEFAULT_COMPANY.id;
  const isInbox = task.projectId === "inbox";
  const isCompleted = Boolean(task.completed);

  if (state.activeCompanyId !== companyId) {
    setActiveCompany(companyId);
  }

  state.activeAllTab = isCompleted ? "completed" : "created";
  state.activeTaskTab = !isInbox && isCompleted ? "completed" : "active";

  if (state.viewMode !== "list") {
    setViewMode("list");
  }

  if (isInbox) {
    setActiveView("view", "workspace");
  } else {
    setActiveView("project", task.projectId);
  }

  window.requestAnimationFrame(() => {
    focusTaskRow(taskId);
    openTaskDialog(taskId);
  });
};

const commitTaskDialogChanges = () => {
  if (!state.editingTaskId || !elements.dialogForm) return false;
  const data = new FormData(elements.dialogForm);
  const title = normaliseTitle(data.get("title") ?? "");
  if (!title) {
    window.alert("Title is required.");
    return false;
  }

  const projectId = data.get("project") || "inbox";
  const sectionId = data.get("section") || getDefaultSectionId(projectId);

  try {
    updateTask(state.editingTaskId, {
      title,
      description: (data.get("description") ?? "").trim(),
      dueDate: data.get("dueDate") || "",
      priority: data.get("priority") || "medium",
      projectId,
      sectionId,
      departmentId: data.get("department") || "",
      assigneeId: data.get("assignee") || "",
      attachments: cloneAttachments(state.dialogAttachmentDraft),
      links,
      completed: elements.dialogForm.completed.checked,
    });
    return true;
  } catch (error) {
    console.error("Failed to update task.", error);
    window.alert("Unable to save changes. Please try again.");
    return false;
  }
};

const autoSaveTaskDialog = () => {
  if (!elements.taskDialog) return false;
  if (!state.editingTaskId) {
    closeTaskDialog();
    return true;
  }
  const saved = commitTaskDialogChanges();
  if (saved) {
    closeTaskDialog();
  }
  return saved;
};

const handleDialogSubmit = (event) => {
  event.preventDefault();
  commitTaskDialogChanges() && closeTaskDialog();
};

const handleDialogClick = (event) => {
  const action = event.target.dataset.action;
  if (action === "close") {
    event.preventDefault();
    autoSaveTaskDialog();
  } else if (action === "toggle-task-completion" && state.editingTaskId) {
    const current = state.tasks.find((entry) => entry.id === state.editingTaskId);
    if (!current || current.deletedAt) return;
    const updated = updateTask(state.editingTaskId, { completed: !current.completed });
    if (updated) {
      updateTaskDialogCompletionState(updated);
      focusTaskRow(updated.id);
    }
  } else if (action === "delete" && state.editingTaskId) {
    const confirmed = window.confirm("Delete this task? This cannot be undone.");
    if (confirmed) {
      removeTask(state.editingTaskId);
      closeTaskDialog();
    }
  } else if (action === "dialog-add-link") {
    event.preventDefault();
    if (elements.dialogLinksList) {
      createLinkRow(elements.dialogLinksList);
    }
  } else if (action === "remove-link") {
    const row = event.target.closest('[data-link-row]');
    if (row && elements.dialogLinksList?.contains(row)) {
      event.preventDefault();
      row.remove();
      if (elements.dialogLinksList.childElementCount === 0) {
        createLinkRow(elements.dialogLinksList);
      }
    }
  } else if (action === "remove-dialog-attachment") {
    const attachmentId = event.target.dataset.attachmentId;
    if (attachmentId) {
      removeDialogAttachment(attachmentId);
    }
  }
};

const handleDialogProjectChange = () => {
  const projectId = elements.dialogForm.elements.project.value;
  populateSectionOptions(elements.dialogForm.elements.section, projectId);
};

const handleDialogDepartmentChange = () => {
  const departmentId = elements.dialogForm.elements.department.value;
  populateMemberOptions(elements.dialogForm.elements.assignee, "", departmentId);
};

const handleQuickAddProjectChange = () => {
  const projectId = elements.quickAddProject.value;
  populateSectionOptions(elements.quickAddSection, projectId);
};

const handleQuickAddDepartmentChange = () => {
  const departmentId = elements.quickAddDepartment.value;
  populateMemberOptions(elements.quickAddAssignee, "", departmentId);
};

const handleMeetingDepartmentChange = () => {
  const departmentId = elements.meetingDepartment ? elements.meetingDepartment.value : "";
  populateMemberOptions(elements.meetingAssignee, "", departmentId);
};

const handleEmailDepartmentChange = () => {
  const departmentId = elements.emailDepartment ? elements.emailDepartment.value : "";
  populateMemberOptions(elements.emailAssignee, "", departmentId);
};

const handleAddSection = () => {
  if (state.activeView.type !== "project") {
    window.alert("Select a project before adding sections.");
    return;
  }
  const name = window.prompt("Section name");
  if (!name) return;
  try {
    createSection(state.activeView.value, name);
    saveSections();
    render();
  } catch (error) {
    console.error("Failed to add section.", error);
  }
};

const promptCreateCompany = () => {
  const name = window.prompt("Company name");
  if (!name) return;
  const company = createCompany(name);
  if (!company) return;
  setActiveCompany(company.id);
  const projectName = window.prompt("Add a first project for this company?");
  if (projectName) {
    createProject(projectName, company.id);
  } else {
    renderProjectDropdown();
  }
};

const prepareCompanyDialog = (companyId) => {
  if (!elements.companyForm) return false;
  const company = getCompanyById(companyId);
  if (!company) return false;
  const input = elements.companyNameInput ?? elements.companyForm.elements.companyName;
  if (input) {
    input.value = company.name;
  }
  if (elements.companyError) {
    elements.companyError.textContent = "";
  }
  const subtitle = elements.companyForm.querySelector("[data-company-subtitle]");
  if (subtitle) {
    subtitle.textContent = company.isDefault ? "Default company" : "";
    subtitle.hidden = !company.isDefault;
  }
  const deleteButton = elements.companyForm.querySelector('[data-action="delete-company"]');
  if (deleteButton) {
    deleteButton.hidden = Boolean(company.isDefault);
    deleteButton.disabled = Boolean(company.isDefault);
  }
  return true;
};

const openCompanyDialog = (companyId) => {
  if (!elements.companyDialog) return;
  const prepared = prepareCompanyDialog(companyId);
  if (!prepared) return;
  state.editingCompanyId = companyId;
  if (typeof elements.companyDialog.showModal === "function") {
    elements.companyDialog.showModal();
  } else {
    elements.companyDialog.setAttribute("open", "true");
  }
  window.requestAnimationFrame(() => {
    const input = elements.companyNameInput ?? elements.companyForm?.elements.companyName;
    input?.focus();
    input?.select?.();
  });
};

const closeCompanyDialog = () => {
  state.editingCompanyId = null;
  if (elements.companyForm) {
    elements.companyForm.reset();
  }
  if (elements.companyError) {
    elements.companyError.textContent = "";
  }
  if (!elements.companyDialog) return;
  if (typeof elements.companyDialog.close === "function") {
    elements.companyDialog.close();
  } else {
    elements.companyDialog.removeAttribute("open");
  }
};

const handleCompanyFormSubmit = (event) => {
  event.preventDefault();
  if (!elements.companyForm) return;
  const companyId = state.editingCompanyId;
  const company = getCompanyById(companyId);
  if (!company) {
    closeCompanyDialog();
    return;
  }
  const input = elements.companyNameInput ?? elements.companyForm.elements.companyName;
  const nextName = input?.value?.trim() ?? "";
  if (!nextName) {
    if (elements.companyError) {
      elements.companyError.textContent = "Company name is required.";
    }
    input?.focus();
    return;
  }
  const duplicate = state.companies.some(
    (entry) => entry.id !== companyId && entry.name.toLowerCase() === nextName.toLowerCase()
  );
  if (duplicate) {
    if (elements.companyError) {
      elements.companyError.textContent = "Another company already uses that name.";
    }
    input?.focus();
    input?.select?.();
    return;
  }

  if (nextName !== company.name) {
    company.name = nextName;
    company.updatedAt = new Date().toISOString();
    saveCompanies();
    renderCompanyTabs();
    renderProjectDropdown();
    renderHeader();
  }

  closeCompanyDialog();
};

const handleCompanyFormClick = (event) => {
  const action = event.target.dataset.action;
  if (!action) return;
  if (action === "close-company") {
    event.preventDefault();
    closeCompanyDialog();
    return;
  }
  if (action === "import-company-whatsapp") {
    event.preventDefault();
    const companyId = state.editingCompanyId;
    if (!companyId) return;
    const company = getCompanyById(companyId);
    if (!company) return;
    closeCompanyDialog();
    if (state.activeCompanyId !== companyId) {
      setActiveCompany(companyId);
    }
    window.requestAnimationFrame(() => {
      openWhatsappDialog();
    });
    return;
  }
  if (action === "delete-company") {
    event.preventDefault();
    const companyId = state.editingCompanyId;
    if (!companyId) return;
    const deleted = deleteCompany(companyId);
    if (deleted) {
      closeCompanyDialog();
    } else if (elements.companyError) {
      elements.companyError.textContent = "Unable to delete this company.";
    }
  }
};

const handleCompanyTabsClick = (event) => {
  const menuButton = event.target.closest('[data-action="open-company-dialog"]');
  if (menuButton) {
    event.preventDefault();
    openCompanyDialog(menuButton.dataset.companyId);
    return;
  }
  const tabButton = event.target.closest("button[data-company-tab]");
  if (tabButton) {
    event.preventDefault();
    setActiveCompany(tabButton.dataset.companyTab);
    hideSearchResults();
  }
};

const handleTaskTabListClick = (event) => {
  const button = event.target.closest("button[data-task-tab]");
  if (!button) return;
  event.preventDefault();
  const tabId = button.dataset.taskTab;
  if (state.activeCompanyId === ALL_COMPANY_ID) {
    if (tabId === state.activeAllTab) return;
    state.activeAllTab = tabId;
  } else {
    if (tabId === state.activeTaskTab) return;
    state.activeTaskTab = tabId;
  }
  savePreferences();
  renderTasks();
};

const handleGlobalActionClick = (event) => {
  const addCompany = event.target.closest('[data-action="add-company"]');
  if (addCompany) {
    event.preventDefault();
    promptCreateCompany();
  }
};

const handleSearchResultClick = (event) => {
  const button = event.target.closest('[data-action="search-select-task"]');
  if (!button) return;
  event.preventDefault();
  navigateToTask(button.dataset.taskId);
  setSearchTerm("");
};

const handleActivityClick = (event) => {
  const moreButton = event.target.closest('[data-action="show-more-activity"]');
  if (moreButton) {
    state.recentActivityLimit = (state.recentActivityLimit || 10) + 10;
    renderActivityFeed();
    return;
  }
  const button = event.target.closest("[data-task-link]");
  if (!button) return;
  event.preventDefault();
  navigateToTask(button.dataset.taskLink);
};

const openProjectEditor = (projectId) => {
  openProjectDialog(projectId);
};

const openSectionEditor = (sectionId) => {
  const section = getSectionById(sectionId);
  if (!section) return;
  const nextName = window.prompt("Rename section", section.name);
  if (nextName && nextName.trim() && nextName.trim() !== section.name) {
    renameSection(section.id, nextName);
  }
  const sections = getSectionsForProject(section.projectId);
  if (sections.length <= 1) return;
  const confirmDelete = window.confirm(
    `Delete the section "${section.name}"? Tasks inside will move to the first remaining section.`
  );
  if (confirmDelete) {
    deleteSection(section.id);
  }
};

const handleProjectMenuClick = (event) => {
  const addButton = event.target.closest('[data-action="add-project-inline"]');
  if (addButton) {
    if (!state.activeCompanyId) {
      window.alert("Select a company before creating projects.");
      return;
    }
    const name = window.prompt("Project name");
    if (name) {
      createProject(name, state.activeCompanyId);
    }
    closeDropdown("project");
    return;
  }

  const editButton = event.target.closest('[data-action="edit-project"]');
  if (editButton) {
    openProjectEditor(editButton.dataset.projectId);
    closeDropdown("project");
    return;
  }

  const meetingButton = event.target.closest('[data-action="quick-meeting"]');
  if (meetingButton) {
    openMeetingDialog(meetingButton.dataset.projectId);
    closeDropdown("project");
    return;
  }

  const emailButton = event.target.closest('[data-action="quick-email"]');
  if (emailButton) {
    openEmailDialog(emailButton.dataset.projectId);
    closeDropdown("project");
    return;
  }

  const selectButton = event.target.closest('button[data-select="project"]');
  if (selectButton) {
    setActiveView("project", selectButton.dataset.projectId);
    closeDropdown("project");
  }
};

const handleUserguideEditToggle = () => {
  state.isEditingUserguide = true;
  renderUserguidePanel();
  elements.userguideEditor?.focus();
};

const handleUserguideCancelEdit = () => {
  state.isEditingUserguide = false;
  renderUserguidePanel();
};

const handleUserguideSave = (event) => {
  event.preventDefault();
  const raw = elements.userguideEditor?.value ?? "";
  const entries = raw
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
  state.userguide = entries.length ? entries : [...DEFAULT_USERGUIDE];
  state.isEditingUserguide = false;
  saveUserguide();
  renderUserguidePanel();
};

const handleBoardClick = (event) => {
  const editButton = event.target.closest('[data-action="edit-section"]');
  if (editButton) {
    event.preventDefault();
    openSectionEditor(editButton.dataset.sectionId);
  }
};

const dropdownElements = {
  project: () => ({
    toggle: elements.projectDropdownToggle,
    menu: elements.projectDropdownMenu,
  }),
};

const setDropdownState = (type, open) => {
  const refs = dropdownElements[type]?.();
  if (!refs) return;
  if (!refs.toggle || !refs.menu) return;
  if (open) {
    closeDropdown(state.openDropdown);
    refs.menu.removeAttribute("hidden");
    refs.toggle.setAttribute("aria-expanded", "true");
    state.openDropdown = type;
  } else {
    refs.menu.setAttribute("hidden", "");
    refs.toggle.setAttribute("aria-expanded", "false");
    if (state.openDropdown === type) {
      state.openDropdown = null;
    }
  }
};

const toggleDropdown = (type) => {
  const isOpen = state.openDropdown === type;
  setDropdownState(type, !isOpen);
};

const closeDropdown = (type) => {
  if (!type) return;
  setDropdownState(type, false);
};

const closeAllDropdowns = () => {
  closeDropdown("company");
  closeDropdown("project");
};

const openUserguide = () => {
  if (!elements.userguidePanel || state.isUserguideOpen) return;
  elements.userguidePanel.hidden = false;
  state.isUserguideOpen = true;
  closeAllDropdowns();
  renderUserguidePanel();
  elements.userguidePanel.scrollIntoView({ behavior: "smooth", block: "start" });
};

const closeUserguide = () => {
  if (!elements.userguidePanel || !state.isUserguideOpen) return;
  elements.userguidePanel.hidden = true;
  state.isUserguideOpen = false;
  if (state.isEditingUserguide) {
    state.isEditingUserguide = false;
    renderUserguidePanel();
  }
};

const toggleUserguide = () => {
  if (state.isUserguideOpen) {
    closeUserguide();
  } else {
    openUserguide();
  }
};

const MS_IN_DAY = 24 * 60 * 60 * 1000;
const WHATSAPP_MESSAGE_REGEX =
  /^(\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4}),?\s+(\d{1,2}:\d{2}(?:\s?[APap][Mm])?)\s+-\s+([^:]+?):\s(.*)$/;

const normaliseYear = (year) => {
  if (!year) return "";
  return year.length === 2 ? `20${year}` : year;
};

const normaliseTimePart = (timePart = "") =>
  timePart
    .replace(/\u202f/g, " ")
    .replace(/\./g, ":")
    .replace(/([APap])\.?M\.?/g, "$1M")
    .toUpperCase();

const parseWhatsappTimestamp = (datePart, timePart) => {
  const cleanTime = normaliseTimePart(timePart);
  const separators = /[\/\.\-]/;
  const parts = datePart.split(separators).map((part) => part.trim());
  if (parts.length < 3) {
    const parsed = Date.parse(`${datePart} ${cleanTime}`);
    return Number.isNaN(parsed) ? null : new Date(parsed);
  }
  let [a, b, c] = parts;
  c = normaliseYear(c);

  const tryParse = (month, day) => {
    const candidate = Date.parse(`${month}/${day}/${c} ${cleanTime}`);
    return Number.isNaN(candidate) ? null : new Date(candidate);
  };

  let parsed = tryParse(a, b);
  if (parsed) return parsed;

  parsed = tryParse(b, a);
  if (parsed) return parsed;

  const fallback = Date.parse(`${datePart} ${cleanTime}`);
  return Number.isNaN(fallback) ? null : new Date(fallback);
};

const extractChatName = (line, fallback) => {
  if (!line) return fallback;
  const trimmed = line.replace(/^\uFEFF/, "").trim();
  const lower = trimmed.toLowerCase();
  if (lower.startsWith("whatsapp chat with")) {
    return trimmed.split("with").slice(1).join("with").trim() || fallback;
  }
  if (lower.startsWith("messages to this chat")) return fallback;
  if (lower.startsWith("chat history with")) {
    return trimmed.split("with").slice(1).join("with").trim() || fallback;
  }
  return fallback;
};

const parseWhatsappTranscript = (rawText, fileName) => {
  const text = rawText.replace(/^\uFEFF/, "");
  const lines = text.split(/\r?\n/);
  const fallbackName = (fileName || "WhatsApp Chat").replace(/\.txt$/i, "").trim() || "WhatsApp Chat";
  let chatName = fallbackName;
  const messages = [];
  let current = null;

  while (lines.length && !lines[0].trim()) {
    lines.shift();
  }
  if (lines.length) {
    chatName = extractChatName(lines[0], fallbackName);
    if (lines[0].toLowerCase().includes("whatsapp chat with")) {
      lines.shift();
    }
  }

  for (const rawLine of lines) {
    if (!rawLine) continue;
    const line = rawLine.replace(/\u202f/g, " ").trim();
    if (!line) continue;
    if (/Messages and calls are end-to-end encrypted/i.test(line)) continue;
    if (/^<Media omitted>$/i.test(line)) continue;

    const match = WHATSAPP_MESSAGE_REGEX.exec(line);
    if (match) {
      const [, datePart, timePart, senderPart, messagePart] = match;
      const timestamp = parseWhatsappTimestamp(datePart, timePart);
      if (!timestamp) continue;
      current = {
        timestamp,
        sender: senderPart.trim(),
        text: messagePart.trim(),
      };
      messages.push(current);
      continue;
    }
    if (current) {
      current.text = `${current.text}\n${rawLine.trim()}`;
    }
  }

  return { chatName, messages };
};

const readWhatsappExport = async (file) => {
  const name = file.name || "whatsapp-export";
  if (name.toLowerCase().endsWith(".zip")) {
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);
    const txtEntry = Object.values(zip.files).find((entry) => entry.name.toLowerCase().endsWith(".txt"));
    if (!txtEntry) {
      throw new Error("The ZIP file does not contain a chat text export.");
    }
    const text = await txtEntry.async("string");
    return parseWhatsappTranscript(text, txtEntry.name);
  }
  const text = await file.text();
  return parseWhatsappTranscript(text, name);
};

const formatDateRange = (start, end) => {
  if (!start || !end) return "";
  const sameDay = start.toDateString() === end.toDateString();
  const formatter = new Intl.DateTimeFormat(undefined, {
    month: "short",
    day: "numeric",
    year: start.getFullYear() !== end.getFullYear() || start.getFullYear() !== new Date().getFullYear() ? "numeric" : undefined,
  });
  if (sameDay) return formatter.format(end);
  return `${formatter.format(start)} - ${formatter.format(end)}`;
};

const buildTranscript = (messages) =>
  messages
    .map(
      (message, index) =>
        `[${index}] ${message.timestamp.toISOString()} | ${message.sender}: ${message.text.replace(/\s+/g, " ").trim()}`,
    )
    .join("\n");

const normalisePriority = (value) => {
  const allowed = ["critical", "very-high", "high", "medium", "low", "optional"];
  if (typeof value !== "string") return "medium";
  const normalised = value.trim().toLowerCase();
  return allowed.includes(normalised) ? normalised : "medium";
};

const normaliseDueDate = (value) => {
  if (!value || typeof value !== "string") return "";
  const trimmed = value.trim();
  if (!trimmed) return "";
  const parsed = Date.parse(trimmed);
  if (Number.isNaN(parsed)) return "";
  const iso = new Date(parsed).toISOString();
  return iso.slice(0, 10);
};

const EST_DATE_FORMATTER = new Intl.DateTimeFormat("en-US", {
  timeZone: "America/New_York",
  month: "short",
  day: "numeric",
  year: "numeric",
});
const EST_TIME_FORMATTER = new Intl.DateTimeFormat("en-US", {
  timeZone: "America/New_York",
  hour: "numeric",
  minute: "2-digit",
  hour12: true,
});

const formatEstDateTime = (date) => {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  const datePart = EST_DATE_FORMATTER.format(date);
  const timePart = EST_TIME_FORMATTER.format(date);
  return `${datePart} at ${timePart} ET`;
};

const simplifyGroupChatName = (name = "") => {
  if (!name) return "";
  const stripped = name.replace(/^WhatsApp Chat with\s+/i, "").trim();
  return stripped || name.trim();
};

const truncateExcerpt = (text = "") => {
  const cleaned = text.replace(/\s+/g, " ").trim();
  if (!cleaned) return "";
  if (cleaned.length <= 280) return cleaned;
  return `${cleaned.slice(0, 277).trimEnd()}...`;
};

const findMessageExcerpt = (messages, timestamp) => {
  if (!(timestamp instanceof Date) || Number.isNaN(timestamp.getTime())) return "";
  const target = timestamp.getTime();
  const match = messages.find(
    (entry) => entry.timestamp && Math.abs(entry.timestamp.getTime() - target) <= 1000,
  );
  return truncateExcerpt(match?.text ?? "");
};

const findCompanyByName = (name) => {
  if (!name || typeof name !== "string") return null;
  const target = name.trim().toLowerCase();
  if (!target) return null;
  return (
    state.companies.find((entry) => entry.name?.trim().toLowerCase() === target) ?? null
  );
};

const ensureWhatsappProjectForCompany = (companyId) => {
  if (!companyId) return null;
  const desiredName = WHATSAPP_PROJECT_NAME.trim();
  const target = desiredName.toLowerCase();
  let project =
    state.projects.find(
      (entry) => entry.companyId === companyId && entry.name?.trim().toLowerCase() === target,
    ) ?? null;
  if (project) return project;

  const timestamp = new Date().toISOString();
  project = {
    id: generateId("project"),
    name: desiredName || "WhatsApp Tasks",
    companyId,
    color: pickProjectColor(state.projects.length),
    createdAt: timestamp,
    updatedAt: timestamp,
  };

  state.projects.push(project);
  saveProjects();
  ensureSectionForProject(project.id);
  return project;
};

const ensureWhatsappSectionForProject = (projectId) => {
  if (!projectId) return null;
  const desiredName = WHATSAPP_SECTION_NAME.trim();
  if (desiredName) {
    const target = desiredName.toLowerCase();
    const existing = getSectionsForProject(projectId).find(
      (section) => section.name?.trim().toLowerCase() === target,
    );
    if (existing) return existing;
    const created = createSection(projectId, desiredName);
    if (created) return created;
  }
  return ensureSectionForProject(projectId);
};

const getWhatsappDestination = () => {
  const jobCompanyId =
    state.importJob?.companyId && state.importJob.companyId !== ALL_COMPANY_ID
      ? state.importJob.companyId
      : null;
  let targetCompanyId =
    jobCompanyId ||
    (state.activeCompanyId && state.activeCompanyId !== ALL_COMPANY_ID
      ? state.activeCompanyId
      : null);

  if (!targetCompanyId && WHATSAPP_COMPANY_NAME.trim()) {
    const fallback = findCompanyByName(WHATSAPP_COMPANY_NAME);
    if (fallback) {
      targetCompanyId = fallback.id;
    }
  }

  if (!targetCompanyId) {
    throw new Error("Open the importer from a specific company to choose where WhatsApp tasks go.");
  }

  const company = getCompanyById(targetCompanyId);
  if (!company) {
    throw new Error("We couldn't find that company. Refresh and try again.");
  }

  let project =
    state.importJob.projectId && getProjectById(state.importJob.projectId)
      ? getProjectById(state.importJob.projectId)
      : null;
  if (!project || project.companyId !== company.id) {
    project = ensureWhatsappProjectForCompany(company.id);
  }
  if (!project) {
    throw new Error("We couldn't prepare a project for WhatsApp tasks. Try again.");
  }
  const section = ensureWhatsappSectionForProject(project.id);
  if (!section) {
    throw new Error("We couldn't prepare a section for WhatsApp tasks. Try again.");
  }

  state.importJob.companyId = company.id;
  state.importJob.projectId = project.id;

  return { company, project, section };
};

const ensureWhatsappLookbackWindow = () => {
  if (Number.isNaN(WHATSAPP_LOOKBACK_DAYS) || WHATSAPP_LOOKBACK_DAYS <= 0) {
    return 30;
  }
  return WHATSAPP_LOOKBACK_DAYS;
};

const looksLikeTask = (candidate) =>
  candidate &&
  typeof candidate === "object" &&
  ("title" in candidate || "description" in candidate || "sourceTimestamp" in candidate);

const safeParseJson = (rawText) => {
  const text = typeof rawText === "string" ? rawText.trim() : "";
  if (!text) return null;
  const attempt = (value) => {
    if (!value) return null;
    try {
      return JSON.parse(value);
    } catch {
      return null;
    }
  };
  let parsed = attempt(text);
  if (parsed) return parsed;
  const firstCurly = text.indexOf("{");
  const lastCurly = text.lastIndexOf("}");
  if (firstCurly !== -1 && lastCurly > firstCurly) {
    parsed = attempt(text.slice(firstCurly, lastCurly + 1));
    if (parsed) return parsed;
  }
  const firstSquare = text.indexOf("[");
  const lastSquare = text.lastIndexOf("]");
  if (firstSquare !== -1 && lastSquare > firstSquare) {
    parsed = attempt(text.slice(firstSquare, lastSquare + 1));
    if (parsed) return parsed;
  }
  return null;
};

const parseActionItemsJson = (payload) => {
  if (Array.isArray(payload)) return payload;
  if (looksLikeTask(payload)) return [payload];
  if (payload && typeof payload === "object") {
    if (Array.isArray(payload.tasks)) return payload.tasks;
    if (looksLikeTask(payload.tasks)) return [payload.tasks];
    if (Array.isArray(payload.items)) return payload.items;
    if (looksLikeTask(payload.items)) return [payload.items];
    if (Array.isArray(payload.actions)) return payload.actions;
    if (looksLikeTask(payload.actions)) return [payload.actions];
  }
  return [];
};

const callOpenRouterForActionItems = async ({
  chatName,
  transcript,
  allowedAssignees,
  model = OPENROUTER_MODEL,
}) => {
  if (!OPENROUTER_API_KEY) {
    throw new Error("Set VITE_OPENROUTER_API_KEY to enable WhatsApp imports.");
  }
  if (!transcript) {
    return { items: [], providerUsed: WHATSAPP_DEFAULT_PROVIDER, modelUsed: model };
  }

  const allowedNames = allowedAssignees.map((entry) => entry.name).filter(Boolean);
  const instructions = `
You are an AI assistant analyzing WhatsApp group conversations for actionable tasks.

Chat: ${chatName}
Transcript lines already fall inside the allowed import window (from the stored start date through now). Read every message in order exactly as provided.

Extraction rules
1. Capture an action item whenever someone requests, commits to, promises, or implies concrete follow-up work. Include direct asks, self-assigned tasks, reminders, and decisions that describe the next step. Skip greetings or pure status updates with no next action.
2. Use real names wherever possible. Strip @ prefixes, and only fall back to a phone number wrapped in single quotes (e.g., '1234567890') when no name exists anywhere in the chat.
3. If several people share responsibility, list every name in the assignee field. If nobody is clearly responsible, set assignee to null.
4. Convert relative timing (tomorrow, Friday, next week) into absolute YYYY-MM-DD dates in America/New_York. When no timing exists, set dueDate to null.
5. Assign priority using exactly one of: "optional", "low", "medium", "high", "very-high", "critical" based on urgency, deadlines, or consequences of delay.
6. The "description" field must be one or two sentences summarizing what needs to happen. Do not copy the raw chat line verbatim; keep it concise.

Allowed assignees: ${allowedNames.length ? allowedNames.join(", ") : "(none)"}.

Output format
Return a pure JSON array (no code fences). Always return an array, even if it contains only one task. Example:

[
  {
    "title": "Example task",
    "description": "Summarize the decision and next step in one or two sentences.",
    "assignee": "Name",
    "dueDate": "2025-10-31",
    "priority": "medium",
    "sourceTimestamp": "2025-10-31T12:00:00.000Z",
    "sourceSender": "Jordan Taylor"
  }
]

Each object must contain exactly:

{
  "title": "Short imperative task summary",
  "description": "One or two sentences summarizing the required work",
  "assignee": "Comma-separated real names or null",
  "dueDate": "YYYY-MM-DD or null",
  "priority": "optional|low|medium|high|very-high|critical",
  "sourceTimestamp": "ISO timestamp from the triggering message",
  "sourceSender": "Name of the message author"
}

If no qualifying action items exist, return [].
`.trim();

  const messages = [
    { role: "system", content: instructions },
    { role: "user", content: `Transcript:\n${transcript}` },
  ];

  const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${OPENROUTER_API_KEY}`,
      "HTTP-Referer": typeof window !== 'undefined' ? window.location.origin : 'https://openrouter.ai',
      "X-Title": "Synergy Tasks",
    },
    body: JSON.stringify({
      model,
      messages,
      temperature: 0.2,
      max_tokens: 1200,
      response_format: { type: "json_object" },
    }),
  });

  const payload = await response.json();
  if (!response.ok) {
    const message = payload?.error?.message || "OpenRouter request failed.";
    throw new Error(message);
  }

  let content = payload?.choices?.[0]?.message?.content ?? "";
  content = content.trim();
  if (!content) {
    return {
      items: [],
      providerUsed: payload?.provider ?? WHATSAPP_DEFAULT_PROVIDER,
      modelUsed: payload?.model ?? model,
    };
  }

  if (/^```/i.test(content)) {
    content = content.replace(/^```(?:json)?\s*/i, "").replace(/```$/i, "").trim();
  }

  const parsed = safeParseJson(content);
  if (!parsed) {
    console.error("Failed to parse OpenRouter response", content);
    throw new Error("OpenRouter returned an unexpected response.");
  }

  return {
    items: parseActionItemsJson(parsed),
    providerUsed: payload?.provider ?? WHATSAPP_DEFAULT_PROVIDER,
    modelUsed: payload?.model ?? model,
  };
};

const summariseImportStats = ({ chatName, messageCount, taskCount, model, provider }) => {
  const lines = [];
  lines.push(`${messageCount} new message${messageCount === 1 ? "" : "s"} analysed`);
  lines.push(`${taskCount} action item${taskCount === 1 ? "" : "s"} created`);
  lines.push(`Source chat: ${chatName}`);
  if (model) {
    const suffix = provider ? ` (${provider})` : "";
    lines.push(`Model: ${model}${suffix}`);
  }
  return lines;
};

const logWhatsappActionsToSheet = async ({ chatName, tasks }) => {
  if (!WHATSAPP_LOG_SHEET_ID) return;
  // Placeholder for future spreadsheet logging.
  console.info(
    `[WhatsApp Import] Spreadsheet logging not yet implemented. Pending ${tasks.length} item(s) for ${chatName}.`,
  );
};

const resolveDefaultWhatsappProject = (companyId) => {
  if (!companyId) return null;
  ensureWhatsappProjectForCompany(companyId);
  const projects = getProjectsForCompany(companyId);
  return projects[0]?.id ?? null;
};

const resetWhatsappImport = (companyId = state.activeCompanyId) => {
  const model = state.importJob.model || OPENROUTER_MODEL;
  const provider = state.importJob.provider || WHATSAPP_DEFAULT_PROVIDER;
  const previousCompanyId =
    state.importJob.companyId && state.importJob.companyId !== ALL_COMPANY_ID
      ? state.importJob.companyId
      : null;
  const candidateCompanyId =
    companyId && companyId !== ALL_COMPANY_ID ? companyId : null;
  const resolvedCompanyId = candidateCompanyId ?? previousCompanyId ?? null;
  let projectId = state.importJob.projectId || null;
  if (resolvedCompanyId) {
    ensureWhatsappProjectForCompany(resolvedCompanyId);
    const projects = getProjectsForCompany(resolvedCompanyId);
    const remembered = state.companyRecents?.[resolvedCompanyId];
    if (!projectId || !projects.some((project) => project.id === projectId)) {
      if (remembered && projects.some((project) => project.id === remembered)) {
        projectId = remembered;
      } else {
        projectId = resolveDefaultWhatsappProject(resolvedCompanyId);
      }
    }
  } else {
    projectId = null;
  }
  state.importJob = {
    file: null,
    status: "idle",
    error: "",
    stats: null,
    model,
    provider,
    companyId: resolvedCompanyId,
    projectId,
  };
  if (elements.whatsappForm) {
    elements.whatsappForm.reset();
  }
  if (elements.whatsappModel) {
    elements.whatsappModel.value = model;
  }
  renderWhatsappImport();
};

const renderWhatsappImport = () => {
  const { file, status, error, stats, model, companyId } = state.importJob;
  if (elements.whatsappPreview) {
    if (stats && file) {
      elements.whatsappPreview.hidden = false;
      if (elements.whatsappFileLabel) {
        elements.whatsappFileLabel.textContent = file.name;
      }
      if (elements.whatsappRangeLabel) {
        elements.whatsappRangeLabel.textContent = stats.range
          ? `Action window: ${stats.range}`
          : "Ready to analyse";
      }
      if (elements.whatsappSummary) {
        elements.whatsappSummary.replaceChildren(
          ...(Array.isArray(stats.summary) && stats.summary.length
            ? stats.summary.map((line) => {
                const item = document.createElement("li");
                item.textContent = line;
                return item;
              })
            : []),
        );
        elements.whatsappSummary.hidden = !(Array.isArray(stats.summary) && stats.summary.length);
      }
    } else {
      elements.whatsappPreview.hidden = true;
    }
  }

  if (elements.whatsappError) {
    if (error) {
      elements.whatsappError.hidden = false;
      elements.whatsappError.textContent = error;
    } else {
      elements.whatsappError.hidden = true;
      elements.whatsappError.textContent = "";
    }
  }

  if (elements.whatsappModel) {
    elements.whatsappModel.value = model;
  }
  if (elements.whatsappCompany) {
    const company = companyId ? getCompanyById(companyId) : null;
    elements.whatsappCompany.textContent = company
      ? `Company: ${company.name}`
      : "Select a company view before importing.";
  }
  if (elements.whatsappProject) {
    const select = elements.whatsappProject;
    if (companyId) {
      ensureWhatsappProjectForCompany(companyId);
      const projects = getProjectsForCompany(companyId);
      if (
        !state.importJob.projectId ||
        !projects.some((project) => project.id === state.importJob.projectId)
      ) {
        state.importJob.projectId = resolveDefaultWhatsappProject(companyId);
      }
      select.replaceChildren(
        ...projects.map((project) => {
          const option = document.createElement("option");
          option.value = project.id;
          option.textContent = project.name;
          return option;
        }),
      );
      select.disabled = status === "processing" || !projects.length;
      if (
        state.importJob.projectId &&
        projects.some((project) => project.id === state.importJob.projectId)
      ) {
        select.value = state.importJob.projectId;
      } else if (projects.length) {
        select.value = projects[0].id;
        state.importJob.projectId = projects[0].id;
      }
    } else {
      select.replaceChildren();
      select.disabled = true;
    }
  }
  if (elements.whatsappFile) {
    elements.whatsappFile.disabled = status === "processing";
  }
  const submitButton = elements.whatsappForm?.querySelector('[data-action="submit-whatsapp"]');
  if (submitButton) {
    submitButton.disabled = status === "processing" || !file;
    submitButton.textContent =
      status === "processing" ? "Processing..." : status === "completed" ? "Run again" : "Fetch action items";
  }
};

const openWhatsappDialog = () => {
  if (!state.activeCompanyId || state.activeCompanyId === ALL_COMPANY_ID) {
    window.alert("Select a company first, then import the WhatsApp chat from that view.");
    return;
  }
  resetWhatsappImport(state.activeCompanyId);
  const dialog = elements.whatsappDialog;
  if (!dialog) return;
  if (typeof dialog.showModal === "function") {
    dialog.showModal();
  } else {
    dialog.setAttribute("open", "");
  }
};

const closeWhatsappDialog = () => {
  const dialog = elements.whatsappDialog;
  if (!dialog) return;
  if (typeof dialog.close === "function") {
    dialog.close();
  } else {
    dialog.removeAttribute("open");
  }
};

const handleWhatsappFileChange = (event) => {
  const file = event.target.files?.[0] ?? null;
  state.importJob.file = file;
  state.importJob.status = "idle";
  state.importJob.error = "";
  state.importJob.stats = file
    ? {
        range: "Ready to analyse",
        summary: [],
      }
    : null;
  renderWhatsappImport();
};

const processWhatsappImport = async () => {
  const file = state.importJob.file;
  if (!file) {
    throw new Error("Select a WhatsApp export before importing.");
  }

  const { company, project, section } = getWhatsappDestination();
  const selectedModel = (state.importJob.model || OPENROUTER_MODEL).trim() || OPENROUTER_MODEL;
  const { chatName, messages } = await readWhatsappExport(file);
  if (!messages.length) {
    throw new Error("No messages were found in the chat export.");
  }

  const lookbackDays = ensureWhatsappLookbackWindow();
  const windowStart = new Date(Date.now() - lookbackDays * MS_IN_DAY);
  const chatIdentifier = chatName || file.name || "default-chat";
  const chatKey = `${company.id}::${chatIdentifier}`;
  const lastProcessedISO =
    state.imports.whatsapp[chatKey] ?? state.imports.whatsapp[chatIdentifier] ?? null;
  const lastProcessedTime = lastProcessedISO ? new Date(lastProcessedISO).getTime() : null;

  const filtered = messages
    .filter((message) => {
      const time = message.timestamp.getTime();
      if (Number.isNaN(time)) return false;
      if (time < windowStart.getTime()) return false;
      if (lastProcessedTime && time <= lastProcessedTime) return false;
      return true;
    })
    .sort((a, b) => a.timestamp.getTime() - b.timestamp.getTime());

  if (!filtered.length) {
    const rangeLabel = formatDateRange(windowStart, new Date());
    const summary = [
      lastProcessedISO
        ? `No new messages since ${new Date(lastProcessedISO).toLocaleString()}`
        : "No recent messages found in the last 30 days",
    ];
    summary.push(`Model: ${selectedModel} (${WHATSAPP_DEFAULT_PROVIDER})`);
    if (rangeLabel) {
      summary.unshift(`Window: ${rangeLabel}`);
    }
    summary.push(`Company: ${company.name}`);
    summary.push(`Project: ${project.name}`);
    state.importJob.model = selectedModel;
    state.importJob.provider = WHATSAPP_DEFAULT_PROVIDER;
    state.importJob.stats = {
      range: rangeLabel,
      summary,
    };
    renderWhatsappImport();
    return;
  }

  const limited = filtered.slice(-Math.max(10, Math.min(filtered.length, MAX_WHATSAPP_LINES || 2000)));
  const transcript = buildTranscript(limited);
  const allowedAssignees = state.members.map((member) => ({ id: member.id, name: member.name }));
  const { items: actionItems, providerUsed, modelUsed } = await callOpenRouterForActionItems({
    chatName,
    transcript,
    allowedAssignees,
    model: selectedModel,
  });
  const effectiveProvider = providerUsed || WHATSAPP_DEFAULT_PROVIDER;
  const effectiveModel = modelUsed || selectedModel;

  const memberLookup = new Map(
    allowedAssignees
      .filter((member) => member.name)
      .map((member) => [member.name.trim().toLowerCase(), member.id]),
  );

  const createdTasks = [];
  const latestTimestamp = filtered[filtered.length - 1].timestamp;
  const earliestTimestamp = filtered[0].timestamp;

  for (const item of actionItems) {
    if (!item || typeof item !== "object") continue;
    const title = String(item.title ?? "").trim();
    if (!title) continue;

    const sourceTimestamp = item.sourceTimestamp ? new Date(item.sourceTimestamp) : null;
    const createdAt =
      sourceTimestamp && !Number.isNaN(sourceTimestamp.getTime())
        ? sourceTimestamp.toISOString()
        : limited[limited.length - 1].timestamp.toISOString();

    const assigneeRaw = typeof item.assignee === "string" ? item.assignee.trim() : "";
    const assigneeValue = assigneeRaw && assigneeRaw.toLowerCase() !== "null" ? assigneeRaw : "";
    const assigneeKey = assigneeValue ? assigneeValue.toLowerCase() : "";
    const assigneeId = assigneeKey && memberLookup.has(assigneeKey) ? memberLookup.get(assigneeKey) : "";

    const dueDate = normaliseDueDate(item.dueDate);
    const priority = normalisePriority(item.priority);
    const sender = typeof item.sourceSender === "string" ? item.sourceSender.trim() : "";
    const detail = typeof item.description === "string" ? item.description.trim() : "";
    const summaryLine = detail || title;
    const excerptText = findMessageExcerpt(limited, sourceTimestamp);
    const excerptLine = excerptText ? `Excerpt: "${excerptText}"` : "";
    const senderTimestampLabel = formatEstDateTime(sourceTimestamp);
    const senderLine =
      sender && senderTimestampLabel
        ? `Sender: ${sender} · ${senderTimestampLabel}`
        : sender
          ? `Sender: ${sender}`
          : senderTimestampLabel
            ? `Sender: ${senderTimestampLabel}`
            : "";
    const groupLine = chatName ? `Group chat: ${simplifyGroupChatName(chatName)}` : "";
    const descriptionSegments = [summaryLine, excerptLine, senderLine, groupLine];

    const task = addTask({
      title,
      description: descriptionSegments.filter(Boolean).join("\n\n"),
      source: "whatsapp",
      projectId: project.id,
      sectionId: section.id,
      assigneeId,
      departmentId: "",
      priority,
      dueDate: dueDate || "",
      createdAt,
    });
    createdTasks.push(task);
  }

  if (createdTasks.length) {
    render();
    updateDashboardMetrics();
    renderActivityFeed();
    await logWhatsappActionsToSheet({ chatName, tasks: createdTasks });
  }

  state.imports.whatsapp[chatKey] = latestTimestamp.toISOString();
  if (chatIdentifier in state.imports.whatsapp && chatIdentifier !== chatKey) {
    delete state.imports.whatsapp[chatIdentifier];
  }
  saveImports();

  state.importJob.model = effectiveModel;
  state.importJob.provider = effectiveProvider;
  const importRange = formatDateRange(earliestTimestamp, latestTimestamp);
  const summaryLines = summariseImportStats({
    chatName,
    messageCount: filtered.length,
    taskCount: createdTasks.length,
    model: effectiveModel,
    provider: effectiveProvider,
  });
  if (importRange) {
    summaryLines.unshift(`Window: ${importRange}`);
  }
  summaryLines.push(`Company: ${company.name}`);
  summaryLines.push(`Project: ${project.name}`);
  state.importJob.stats = {
    range: importRange,
    summary: summaryLines,
  };
  renderWhatsappImport();

  if (!createdTasks.length) {
    window.alert("No action items were found in the latest messages.");
  } else {
    window.alert(`Imported ${createdTasks.length} action item${createdTasks.length === 1 ? "" : "s"} from ${chatName}.`);
  }
};

const handleWhatsappSubmit = async (event) => {
  event.preventDefault();
  if (!state.importJob.file) {
    state.importJob.error = "Select a WhatsApp export before importing.";
    renderWhatsappImport();
    return;
  }
  state.importJob.status = "processing";
  state.importJob.error = "";
  renderWhatsappImport();
  try {
    await processWhatsappImport();
    state.importJob.status = "completed";
  } catch (error) {
    console.error("Failed to import WhatsApp export", error);
    state.importJob.error =
      error?.message ?? "We couldn't process the export. Please try again or check the file.";
    state.importJob.status = "idle";
    renderWhatsappImport();
    return;
  } finally {
    renderWhatsappImport();
  }
};

const handleWhatsappModelChange = (event) => {
  const value = event.target?.value?.trim();
  state.importJob.model = value || OPENROUTER_MODEL;
  renderWhatsappImport();
};

const handleWhatsappProjectChange = (event) => {
  const value = event.target?.value?.trim();
  const companyId = state.importJob.companyId;
  state.importJob.projectId = value || null;
  if (companyId && value) {
    state.companyRecents[companyId] = value;
  }
  renderWhatsappImport();
};

const handleWhatsappDialogClick = (event) => {
  const action = event.target?.dataset?.action;
  if (action === "cancel-whatsapp") {
    event.preventDefault();
    closeWhatsappDialog();
  }
};

const handleGlobalClick = (event) => {
  if (!event.target.closest(".selector-dropdown")) {
    closeAllDropdowns();
  }
  if (!event.target.closest('[data-search-root]')) {
    hideSearchResults();
  }
};

const handleGlobalKeydown = (event) => {
  if (event.key === 'Escape') {
    closeAllDropdowns();
    closeUserguide();
    hideSearchResults();
    if (state.isQuickAddOpen) {
      closeQuickAddForm();
    }
  }
};

const handleStorageEvent = (event) => {
  if (event.key === STORAGE_KEYS.settings) {
    try {
      const nextSettings = event.newValue ? JSON.parse(event.newValue) : defaultSettings();
      state.settings = normaliseSettings(nextSettings);
      applySettings();
      updateDashboardMetrics();
    } catch (error) {
      console.error("Failed to apply updated settings", error);
    }
  }
};

const handleSectionDragStart = (event) => {
  const header = event.currentTarget;
  state.dragSectionId = header.dataset.sectionId;
  header.classList.add('dragging-section');
  event.dataTransfer.effectAllowed = 'move';
};

const handleSectionDragOver = (event) => {
  if (!state.dragSectionId) return;
  event.preventDefault();
  event.dataTransfer.dropEffect = 'move';
  const column = event.currentTarget.closest('.board-column');
  if (!column) return;
  const targetId = column.dataset.sectionId;
  if (!targetId || targetId === state.dragSectionId) return;
  if (state.sectionDropTarget && state.sectionDropTarget !== targetId) {
    const previous = elements.boardColumns?.querySelector(`[data-section-id="${state.sectionDropTarget}"]`);
    previous?.classList.remove('section-drop-target');
  }
  state.sectionDropTarget = targetId;
  column.classList.add('section-drop-target');
};

const handleSectionDragLeave = (event) => {
  const column = event.currentTarget.closest('.board-column');
  column?.classList.remove('section-drop-target');
};

const handleSectionDrop = (event) => {
  event.preventDefault();
  const column = event.currentTarget.closest('.board-column');
  column?.classList.remove('section-drop-target');
  if (!state.dragSectionId || !state.sectionDropTarget) return;
  if (state.dragSectionId === state.sectionDropTarget) return;
  reorderSections(state.dragSectionId, state.sectionDropTarget);
  state.dragSectionId = null;
  state.sectionDropTarget = null;
};

const handleSectionDragEnd = (event) => {
  const header = event.currentTarget;
  header.classList.remove('dragging-section');
  state.dragSectionId = null;
  state.sectionDropTarget = null;
  elements.boardColumns?.querySelectorAll('.section-drop-target').forEach((col) => col.classList.remove('section-drop-target'));
};

const reorderSections = (sourceId, targetId) => {
  if (sourceId === targetId) return;
  const sourceIndex = state.sections.findIndex((section) => section.id === sourceId);
  const targetIndex = state.sections.findIndex((section) => section.id === targetId);
  if (sourceIndex === -1 || targetIndex === -1) return;
  const sourceSection = state.sections[sourceIndex];
  const targetSection = state.sections[targetIndex];
  if (sourceSection.projectId !== targetSection.projectId) return;

  const [moved] = state.sections.splice(sourceIndex, 1);
  const insertIndex = state.sections.findIndex((section) => section.id === targetId);
  state.sections.splice(insertIndex, 0, moved);

  state.sections = state.sections.map((section, index) => ({
    ...section,
    order: index,
  }));
  saveSections();
  render();
};

const handleBoardDragStart = (event) => {
  const card = event.target.closest(".board-card");
  if (!card) return;
  state.dragTaskId = card.dataset.taskId;
  card.classList.add("dragging");
  event.dataTransfer.effectAllowed = "move";
  event.dataTransfer.setData("text/plain", card.dataset.taskId);
};

const handleBoardDragEnd = (event) => {
  const card = event.target.closest(".board-card");
  if (card) card.classList.remove("dragging");
  state.dragTaskId = null;
  elements.boardColumns
    ?.querySelectorAll(".board-column")
    .forEach((column) => column.classList.remove("drag-over"));
};

const handleBoardDragOver = (event) => {
  if (state.dragSectionId) return;
  event.preventDefault();
  const container = event.currentTarget;
  container.classList.add("drag-over");
  container.classList.remove("empty");
  event.dataTransfer.dropEffect = "move";
};

const handleBoardDragLeave = (event) => {
  event.currentTarget.classList.remove("drag-over");
};

const handleBoardDrop = (event) => {
  if (state.dragSectionId) return;
  event.preventDefault();
  const container = event.currentTarget;
  const sectionId = container.dataset.sectionId;
  const taskId = state.dragTaskId || event.dataTransfer.getData("text/plain");
  container.classList.remove("drag-over");
  if (!sectionId || !taskId) return;
  updateTask(taskId, { sectionId });
};

const handleMembersFormSubmit = (event) => {
  event.preventDefault();
  const action = event.submitter?.dataset.action;
  if (action === "add-member") {
    const name = elements.membersForm.memberName.value;
    const departmentId = elements.membersForm.memberDepartment.value;
    try {
      addMember(name, departmentId);
      elements.membersForm.reset();
      elements.membersForm.memberDepartment.value = DEFAULT_DEPARTMENT.id;
    } catch (error) {
      console.error("Failed to add member from form.", error);
    }
  }
};

const handleMembersFormClick = (event) => {
  const action = event.target.dataset.action;
  if (action === "close") {
    closeMembersDialog();
    return;
  }
  if (action === "remove-member") {
    try {
      removeMember(event.target.dataset.memberId);
    } catch (error) {
      console.error("Failed to remove member.", error);
    }
  }
};

const handleDepartmentsFormSubmit = (event) => {
  event.preventDefault();
  const action = event.submitter?.dataset.action;
  if (action === "add-department") {
    const name = elements.departmentsForm.departmentName.value;
    try {
      addDepartment(name);
      elements.departmentsForm.reset();
    } catch (error) {
      console.error("Failed to add department.", error);
    }
  }
};

const handleDepartmentsFormClick = (event) => {
  const action = event.target.dataset.action;
  if (action === "close") {
    closeDepartmentsDialog();
    return;
  }
  if (action === "remove-department") {
    try {
      removeDepartment(event.target.dataset.departmentId);
    } catch (error) {
      console.error("Failed to remove department.", error);
    }
  }
};
const hydrateStateFromLocal = () => {
  state.tasks = loadJSON(STORAGE_KEYS.tasks, []);
  state.projects = loadJSON(STORAGE_KEYS.projects, []);
  state.sections = loadJSON(STORAGE_KEYS.sections, []);
  state.companies = loadJSON(STORAGE_KEYS.companies, [{ ...DEFAULT_COMPANY }]);
  state.members = loadJSON(STORAGE_KEYS.members, []);
  state.departments = loadJSON(STORAGE_KEYS.departments, []);
  state.settings = normaliseSettings(loadJSON(STORAGE_KEYS.settings, defaultSettings()));
  state.userguide = normaliseUserguide(loadJSON(STORAGE_KEYS.userguide, DEFAULT_USERGUIDE));
  const storedImports = loadJSON(STORAGE_KEYS.imports, { whatsapp: {} });
  state.imports = {
    whatsapp: { ...(storedImports?.whatsapp ?? {}) },
  };
  upgradeUserguideIfLegacy();
  ensureUserguideHasLatestEntries();

  ensureDefaultCompany();
  ensureDefaultProject();
  ensureProjectsHaveCompany();
  ensureDefaultDepartment();
  ensureAllProjectsHaveSections();
  state.tasks.forEach(ensureTaskDefaults);

  applyStoredPreferences();
};

const hydrateState = async () => {
  applyStoredPreferences();
  try {
    await startWorkspaceSync();
  } catch (error) {
    console.error('Failed to start Firestore sync, falling back to cached data.', error);
    hydrateStateFromLocal();
    render();
  }
};

const registerEventListeners = () => {
  elements.viewList?.addEventListener("click", handleViewClick);
  elements.boardColumns?.addEventListener("click", handleBoardClick);
  elements.projectDropdownToggle?.addEventListener("click", () => toggleDropdown("project"));
  elements.companyTabs?.addEventListener("click", handleCompanyTabsClick);
  elements.taskTabList?.addEventListener("click", handleTaskTabListClick);
  elements.projectDropdownMenu?.addEventListener("click", handleProjectMenuClick);
  elements.companyForm?.addEventListener("submit", handleCompanyFormSubmit);
  elements.companyForm?.addEventListener("click", handleCompanyFormClick);
  elements.meetingDepartment?.addEventListener("change", handleMeetingDepartmentChange);
  elements.filterMember?.addEventListener("change", handleFilterChange);
  elements.filterDepartment?.addEventListener("change", handleFilterChange);
  elements.filterPriority?.addEventListener("change", handleFilterChange);
  elements.filterDue?.addEventListener("change", handleFilterChange);
  elements.filterBar?.addEventListener("click", handleFilterBarClick);
  elements.companyDialog?.addEventListener("cancel", (event) => {
    event.preventDefault();
    closeCompanyDialog();
  });
  elements.projectForm?.addEventListener("submit", handleProjectFormSubmit);
  elements.projectForm?.addEventListener("click", handleProjectFormClick);
  elements.projectDialog?.addEventListener("cancel", (event) => {
    event.preventDefault();
    closeProjectDialog();
  });
  elements.meetingForm?.addEventListener("submit", handleMeetingFormSubmit);
  elements.meetingForm?.addEventListener("click", handleMeetingFormClick);
  elements.meetingActionList?.addEventListener("change", handleMeetingActionListChange);
  elements.meetingActionList?.addEventListener("input", handleMeetingActionListInput);
  elements.meetingActionInput?.addEventListener("keydown", handleMeetingActionInputKeydown);
  elements.meetingDialog?.addEventListener("cancel", (event) => {
    event.preventDefault();
    autoSaveMeetingDialog();
  });
  elements.emailForm?.addEventListener("submit", handleEmailFormSubmit);
  elements.emailForm?.addEventListener("click", handleEmailFormClick);
  elements.emailDepartment?.addEventListener("change", handleEmailDepartmentChange);
  elements.emailThreadList?.addEventListener("change", handleEmailThreadListChange);
  elements.emailThreadList?.addEventListener("input", handleEmailThreadListInput);
  elements.emailThreadList?.addEventListener("pointerup", handleEmailThreadListPointerUp);
  elements.emailThreadInput?.addEventListener("keydown", handleEmailThreadInputKeydown);
  elements.emailDialog?.addEventListener("cancel", (event) => {
    event.preventDefault();
    autoSaveEmailDialog();
  });
  elements.userguideEditToggle?.addEventListener("click", handleUserguideEditToggle);
  elements.userguideCancelEdit?.addEventListener("click", handleUserguideCancelEdit);
  elements.userguideForm?.addEventListener("submit", handleUserguideSave);
  document
    .querySelectorAll('[data-action="open-userguide"]')
    .forEach((button) => button.addEventListener("click", toggleUserguide));
  document
    .querySelectorAll('[data-action="close-userguide"]')
    .forEach((button) => button.addEventListener("click", closeUserguide));
  document.addEventListener("click", handleGlobalActionClick);
  elements.activityFeed?.addEventListener("click", handleActivityClick);
  elements.importWhatsapp?.addEventListener("click", openWhatsappDialog);
  elements.whatsappForm?.addEventListener("submit", handleWhatsappSubmit);
  elements.whatsappModel?.addEventListener("change", handleWhatsappModelChange);
  elements.whatsappProject?.addEventListener("change", handleWhatsappProjectChange);
  elements.whatsappForm?.addEventListener("click", handleWhatsappDialogClick);
  elements.whatsappFile?.addEventListener("change", handleWhatsappFileChange);
  elements.whatsappDialog?.addEventListener("cancel", (event) => {
    event.preventDefault();
    closeWhatsappDialog();
  });

  elements.taskTabPanels?.addEventListener("click", handleTaskItemOpen);
  elements.taskTabPanels?.addEventListener("click", handleTaskActionClick);

  if (elements.quickAddForm) {
    elements.quickAddForm.addEventListener("submit", handleQuickAddSubmit);
    elements.quickAddForm.addEventListener("click", handleQuickAddClick);
  }
  elements.quickAddCancel?.addEventListener("click", handleQuickAddCancel);
  elements.quickAddProject?.addEventListener("change", handleQuickAddProjectChange);
  elements.quickAddDepartment?.addEventListener("change", handleQuickAddDepartmentChange);
  elements.quickAddAttachments?.addEventListener("change", updateQuickAddAttachmentPreview);
  elements.toggleQuickAdd?.addEventListener("click", toggleQuickAddForm);

  if (elements.searchInput) {
    elements.searchInput.addEventListener("input", handleSearchInput);
  }
  if (elements.searchInputMobile) {
    elements.searchInputMobile.addEventListener("input", (event) => setSearchTerm(event.target.value));
  }
  elements.searchResults?.addEventListener("click", handleSearchResultClick);
  elements.metricFilter?.addEventListener("change", (event) => {
    state.metricsFilter = event.target.value;
    savePreferences();
    updateDashboardMetrics();
  });
  elements.addProject?.addEventListener("click", handleAddProject);
  elements.exportTasks?.addEventListener("click", handleExport);
  elements.addSection?.addEventListener("click", handleAddSection);

  elements.viewToggleButtons.forEach((button) => {
    button.addEventListener("click", () => setViewMode(button.dataset.viewMode));
  });

  if (elements.dialogForm) {
    elements.dialogForm.addEventListener("submit", handleDialogSubmit);
    elements.dialogForm.addEventListener("click", handleDialogClick);
    elements.dialogForm.elements.project.addEventListener("change", handleDialogProjectChange);
    elements.dialogForm.elements.department.addEventListener("change", handleDialogDepartmentChange);
  }
  elements.dialogAttachmentsInput?.addEventListener("change", handleDialogAttachmentsInput);
  elements.taskDialog?.addEventListener("cancel", (event) => {
    event.preventDefault();
    autoSaveTaskDialog();
  });

  elements.manageMembers?.addEventListener("click", openMembersDialog);
  elements.manageDepartments?.addEventListener("click", openDepartmentsDialog);

  elements.membersForm?.addEventListener("submit", handleMembersFormSubmit);
  elements.membersForm?.addEventListener("click", handleMembersFormClick);
  elements.membersDialog?.addEventListener("cancel", (event) => {
    event.preventDefault();
    closeMembersDialog();
  });

  elements.departmentsForm?.addEventListener("submit", handleDepartmentsFormSubmit);
  elements.departmentsForm?.addEventListener("click", handleDepartmentsFormClick);
  elements.departmentsDialog?.addEventListener("cancel", (event) => {
    event.preventDefault();
    closeDepartmentsDialog();
  });

  document.addEventListener("click", handleGlobalClick);
  document.addEventListener("keydown", handleGlobalKeydown);
  window.addEventListener("storage", handleStorageEvent);
};

const init = async () => {
  registerEventListeners();
  resetQuickAddForm();
  await hydrateState();
  initialiseTextareaAutosize();
  renderMeetingActionItems();
  renderEmailThreadMessages();
  render();
};

const startApp = () => {
  init().catch((error) => {
    console.error("Failed to initialise workspace", error);
  });
};

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", startApp);
} else {
  startApp();
}























