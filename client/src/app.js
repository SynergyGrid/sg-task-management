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

const GEMINI_MODEL = import.meta.env.VITE_GEMINI_MODEL || "gemini-1.5-flash";
const GEMINI_API_KEY = import.meta.env.VITE_GEMINI_API_KEY || "";
const WHATSAPP_LOOKBACK_DAYS = Number.parseInt(import.meta.env.VITE_WHATSAPP_LOOKBACK_DAYS ?? "30", 10);
const WHATSAPP_COMPANY_NAME = import.meta.env.VITE_WHATSAPP_COMPANY_NAME || "GENERAL";
const WHATSAPP_PROJECT_NAME = import.meta.env.VITE_WHATSAPP_PROJECT_NAME || "General Project";
const WHATSAPP_DEFAULT_ENDPOINT = (import.meta.env.VITE_WHATSAPP_ENDPOINT || "v1").toLowerCase() === "v1beta" ? "v1beta" : "v1";
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
const DEFAULT_USERGUIDE = [
  "## Navigating the workspace",
  "- Open Workspace to triage unscheduled work across the selected company.",
  "- Use Due Today and Upcoming 7 Days to keep deadlines visible.",
  "- Switch companies from the selector above the project dropdown to scope the board.",
  "## Working inside projects",
  "- Every project opens with Regular Tasks, WhatsApp Tasks, Meetings and Emails.",
  "- Each section surfaces the latest five items. Drag the bottom edge or use See more to review the rest.",
  "- Board view is still available when you need the original sections and drag-and-drop.",
  "## Logging meetings",
  "- Use the Meeting action in the project menu to record session details, attendees and supporting links.",
  "- Meeting entries stay grouped under the project overview and can be edited from the same list.",
  "## Tracking emails",
  "- Use the Email action to capture key threads, their status and reference links.",
  "- Email items surface in the project overview so follow-ups never drift.",
  "## Search and activity",
  "- The search bar now provides live suggestions - select any result to jump straight to the task.",
  "- Recent activity lists the latest updates; use Show more to extend the feed when reviewing history.",
  "## Team administration",
  "- Manage companies, members and departments from Settings so assignments stay accurate.",
  "- Deleting a company or project preserves its tasks under Workspace so nothing is lost.",
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
  editingCompanyId: null,
  editingMeetingId: null,
  editingEmailId: null,
  dragTaskId: null,
  dragSectionId: null,
  sectionDropTarget: null,
  isQuickAddOpen: false,
  isUserguideOpen: false,
  isEditingUserguide: false,
  openDropdown: null,
  dialogAttachmentDraft: [],
  projectSectionLimits: {},
  recentActivityLimit: 10,
  importJob: {
    file: null,
    status: "idle",
    error: "",
    model: GEMINI_MODEL,
    endpoint: WHATSAPP_DEFAULT_ENDPOINT,
    stats: null,
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
  dialogAttachmentsInput: document.querySelector('#dialogForm input[name="attachments"]'),
  dialogAttachmentList: document.querySelector('#dialogForm [data-dialog-attachment-list]'),
  taskTemplate: document.getElementById("taskItemTemplate"),
  activeTasksMetric: document.getElementById("active-tasks"),
  activityFeed: document.getElementById("activity-feed"),
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
  meetingDepartment: document.querySelector('#meetingForm select[name="department"]'),
  meetingPriority: document.querySelector('#meetingForm select[name="priority"]'),
  meetingLinksList: document.querySelector("[data-meeting-links]"),
  meetingError: document.querySelector("[data-meeting-error]"),
  emailDialog: document.getElementById("emailDialog"),
  emailForm: document.getElementById("emailForm"),
  emailProjectLabel: document.querySelector("[data-email-project]"),
  emailDepartment: document.querySelector('#emailForm select[name="department"]'),
  emailPriority: document.querySelector('#emailForm select[name="priority"]'),
  emailStatus: document.querySelector('#emailForm select[name="status"]'),
  emailLinksList: document.querySelector("[data-email-links]"),
  emailError: document.querySelector("[data-email-error]"),
  companyDialog: document.getElementById("companyDialog"),
  companyForm: document.getElementById("companyForm"),
  companyNameInput: document.getElementById("companyName"),
  companyError: document.querySelector("[data-company-error]"),
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
  whatsappEndpoint: document.getElementById("whatsappEndpoint"),
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
});

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
    .filter((task) => !task.deletedAt && matchesSearch(task))
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
  state.tasks.filter((task) => !task.deletedAt && matchesActiveView(task) && matchesSearch(task));

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

const buildMeta = (task, container) => {
  container.textContent = "";

  const due = describeDueDate(task.dueDate);
  if (due.label) pushMetaChip(container, due.label, due.className || "meta-chip");

  if (task.createdAt) {
    const createdDate = new Date(task.createdAt);
    if (!Number.isNaN(createdDate.getTime())) {
      const createdLabel = new Intl.DateTimeFormat(undefined, {
        month: "short",
        day: "numeric",
        year: createdDate.getFullYear() !== new Date().getFullYear() ? "numeric" : undefined,
      }).format(createdDate);
      pushMetaChip(container, `Created ${createdLabel}`, "meta-chip subtle");
    }
  }
  if (task.deletedAt) {
    const deletedDate = new Date(task.deletedAt);
    if (!Number.isNaN(deletedDate.getTime())) {
      const deletedLabel = new Intl.DateTimeFormat(undefined, {
        month: "short",
        day: "numeric",
        year: deletedDate.getFullYear() !== new Date().getFullYear() ? "numeric" : undefined,
      }).format(deletedDate);
      pushMetaChip(container, `Deleted ${deletedLabel}`, "meta-chip danger");
    }
  }

  const section = getSectionById(task.sectionId);
  if (section && !(state.viewMode === "board" && state.activeView.type === "project")) {
    pushMetaChip(container, section.name);
  }
  const priorityLabel = PRIORITY_LABELS[task.priority];
  if (priorityLabel && task.priority !== "medium") {
    pushMetaChip(container, priorityLabel);
  }

  if (shouldShowProjectChip(task)) {
    const project = getProjectById(task.projectId);
    pushMetaChip(container, project ? project.name : "Unknown project");
  }

  const member = getMemberById(task.assigneeId);
  if (member) pushMetaChip(container, member.name);

  const department = getDepartmentById(task.departmentId);
  if (department && !department.isDefault) pushMetaChip(container, department.name);

  if ((task.source ?? "") === "whatsapp") {
    pushMetaChip(container, "WhatsApp import");
  }

  if ((task.kind ?? "task") === "meeting") {
    const meetingType = task.metadata?.meetingType;
    const attendees = task.metadata?.attendees;
    if (meetingType) pushMetaChip(container, `Meeting: ${meetingType}`);
    if (attendees) pushMetaChip(container, `Attendees: ${attendees}`);
    if (task.metadata?.actionItems) pushMetaChip(container, "Action items");
  } else if ((task.kind ?? "task") === "email") {
    const emailAddress = task.metadata?.emailAddress;
    const status = task.metadata?.status;
    if (emailAddress) pushMetaChip(container, emailAddress);
    if (status) pushMetaChip(container, `Status: ${status}`);
  }

  if (Array.isArray(task.links) && task.links.length) {
    pushMetaChip(
      container,
      `${task.links.length} link${task.links.length === 1 ? "" : "s"}`
    );
  }

  const archivedFrom = task.metadata?.archivedFrom;
  if (archivedFrom?.projectName) {
    const archivedLabel = archivedFrom.companyName
      ? `From: ${archivedFrom.projectName} (${archivedFrom.companyName})`
      : `From: ${archivedFrom.projectName}`;
    pushMetaChip(container, archivedLabel);
  }

  if (task.description) pushMetaChip(container, "Notes");
};

const renderTaskItem = (task) => {
  const fragment = elements.taskTemplate.content.cloneNode(true);
  const item = fragment.querySelector(".task-item");
  const checkbox = fragment.querySelector('input[type="checkbox"]');
  const titleEl = fragment.querySelector(".task-title");
  const metaEl = fragment.querySelector(".task-meta");
  const contentEl = fragment.querySelector(".task-content");

  item.dataset.taskId = task.id;
  item.dataset.priority = task.priority;
  titleEl.textContent = task.title;
  titleEl.title = task.description ?? "";
  checkbox.checked = task.completed;
  item.classList.toggle("completed", task.completed);
  if (task.deletedAt) {
    item.classList.add("deleted");
    checkbox.disabled = true;
  } else {
    checkbox.disabled = false;
    item.classList.remove("deleted");
  }

  buildMeta(task, metaEl);

  if (Array.isArray(task.attachments) && task.attachments.length && contentEl) {
    const row = document.createElement("div");
    row.className = "attachments-row";
    task.attachments.forEach((attachment) => {
      row.append(createAttachmentChip(attachment));
    });
    contentEl.append(row);
  }

  if (Array.isArray(task.links) && task.links.length && contentEl) {
    const row = document.createElement("div");
    row.className = "links-row";
    task.links.forEach((link) => {
      const title = link?.title?.trim() || link?.url?.trim() || "Link";
      if (link?.url) {
        const anchor = document.createElement("a");
        anchor.href = link.url;
        anchor.target = "_blank";
        anchor.rel = "noopener";
        anchor.className = "link-chip";
        anchor.textContent = title;
        row.append(anchor);
      } else {
        const span = document.createElement("span");
        span.className = "link-chip disabled";
        span.textContent = title;
        row.append(span);
      }
    });
    contentEl.append(row);
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
      .filter((task) => !task.deletedAt && matchesSearch(task))
      .sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());
  } else if (state.activeAllTab === "updated") {
    scopedTasks = state.tasks
      .filter((task) => !task.deletedAt && matchesSearch(task))
      .sort(
        (a, b) =>
          new Date(b.updatedAt || b.createdAt).getTime() -
          new Date(a.updatedAt || a.createdAt).getTime()
      );
  } else {
    scopedTasks = state.tasks
      .filter((task) => task.deletedAt && matchesSearch(task))
      .sort((a, b) => new Date(b.deletedAt ?? 0).getTime() - new Date(a.deletedAt ?? 0).getTime());
  }

  panel.append(
    renderSimpleTaskList(
      scopedTasks,
      state.activeAllTab === "deleted"
        ? "No deleted tasks yet."
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
  card.addEventListener("dblclick", () => openTaskDialog(task.id));

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
  const { profile, theme } = state.settings;
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
};

const renderSidebar = () => {
  renderCompanyTabs();
  renderProjectDropdown();
  updateActiveNav();
  updateViewCounts();
  renderActivityFeed();
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

const attachTextareaAutosize = (textarea) => {
  if (!textarea) return;
  autoResizeTextarea(textarea);
  textarea.addEventListener("input", () => autoResizeTextarea(textarea));
};

const initialiseTextareaAutosize = () => {
  const quickAddDescription = elements.quickAddForm?.elements?.description;
  attachTextareaAutosize(quickAddDescription);

  if (elements.dialogForm) {
    attachTextareaAutosize(elements.dialogForm.elements.description);
  }

  const meetingForm = elements.meetingForm;
  if (meetingForm) {
    attachTextareaAutosize(meetingForm.elements.attendees);
    attachTextareaAutosize(meetingForm.elements.notes);
  }

  const emailForm = elements.emailForm;
  if (emailForm) {
    attachTextareaAutosize(emailForm.elements.notes);
  }
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
  if (!project) return;
  const trimmed = nextName.trim();
  if (!trimmed) return;
  const exists = state.projects.some(
    (entry) =>
      entry.id !== projectId &&
      entry.companyId === project.companyId &&
      entry.name.toLowerCase() === trimmed.toLowerCase(),
  );
  if (exists) {
    window.alert("Another project in this company already uses that name.");
    return;
  }
  project.name = trimmed;
  project.updatedAt = new Date().toISOString();
  saveProjects();
  renderProjectDropdown();
  renderHeader();
  renderTasks();
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
  return true;
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

const handleTaskCheckboxChange = (event) => {
  if (event.target.type !== "checkbox") return;
  const item = event.target.closest(".task-item");
  if (!item) return;
  updateTask(item.dataset.taskId, { completed: event.target.checked });
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
  const item = button.closest(".task-item");
  if (!item) return;
  const taskId = item.dataset.taskId;
  const task = state.tasks.find((entry) => entry.id === taskId);
  if (!task) return;

  if (action === "edit") {
    const kind = task.kind ?? "task";
    if (kind === "meeting") {
      openMeetingDialog(task.projectId, task);
      return;
    }
    if (kind === "email") {
      openEmailDialog(task.projectId, task);
      return;
    }
    openTaskDialog(taskId);
  } else if (action === "delete") {
    const confirmed = window.confirm("Delete this task? This cannot be undone.");
    if (confirmed) {
      removeTask(taskId);
    }
  }
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
  if (elements.meetingPriority) {
    elements.meetingPriority.value = task?.priority ?? "medium";
  }
  form.elements.date.value = task?.dueDate ?? "";
  form.elements.meetingType.value = task?.metadata?.meetingType ?? "";
  form.elements.attendees.value = task?.metadata?.attendees ?? "";
  form.elements.actionItems.checked = Boolean(task?.metadata?.actionItems);
  form.elements.title.value = task?.title ?? "";
  if (form.elements.notes) {
    form.elements.notes.value = task?.description ?? "";
  }
  resetLinkList(elements.meetingLinksList, Array.isArray(task?.links) ? task.links : []);
  if (elements.meetingError) elements.meetingError.textContent = "";
  autoResizeTextarea(form.elements.attendees);
  autoResizeTextarea(form.elements.notes);
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
  if (form) {
    form.reset();
    resetLinkList(elements.meetingLinksList);
    if (elements.meetingError) elements.meetingError.textContent = "";
  }
  if (dialog) {
    if (typeof dialog.close === "function") dialog.close();
    else dialog.removeAttribute("open");
  }
};

const handleMeetingFormSubmit = (event) => {
  event.preventDefault();
  const form = elements.meetingForm;
  if (!form) return;
  const titleField = form.elements.title;
  const title = titleField?.value?.trim() ?? "";
  if (!title) {
    if (elements.meetingError) {
      elements.meetingError.textContent = "Title is required.";
    }
    titleField?.focus();
    return;
  }

  const projectId = form.elements.projectId.value || "inbox";
  ensureSectionForProject(projectId);
  const sectionId = getDefaultSectionId(projectId);
  const links = collectLinks(elements.meetingLinksList);
  const existing = state.tasks.find((task) => task.id === state.editingMeetingId);
  const metadata = existing?.metadata && typeof existing.metadata === "object" ? { ...existing.metadata } : {};
  metadata.meetingType = form.elements.meetingType.value;
  metadata.attendees = form.elements.attendees.value.trim();
  metadata.actionItems = form.elements.actionItems.checked;
  metadata.links = links;

  const payload = {
    title,
    description: form.elements.notes?.value?.trim() ?? "",
    dueDate: form.elements.date.value || "",
    priority: form.elements.priority.value || "medium",
    projectId,
    sectionId,
    departmentId: form.elements.department.value || "",
    kind: "meeting",
    source: existing?.source ?? "manual",
    metadata,
    links,
  };

  if (state.editingMeetingId && existing) {
    updateTask(state.editingMeetingId, payload);
  } else {
    addTask(payload);
  }

  closeMeetingDialog();
};

const handleMeetingFormClick = (event) => {
  const action = event.target.dataset.action;
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
  if (action === "close-meeting") {
    event.preventDefault();
    closeMeetingDialog();
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
  if (elements.emailError) elements.emailError.textContent = "";
  autoResizeTextarea(form.elements.notes);
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
  if (form) {
    form.reset();
    resetLinkList(elements.emailLinksList);
    if (elements.emailError) elements.emailError.textContent = "";
  }
  if (dialog) {
    if (typeof dialog.close === "function") dialog.close();
    else dialog.removeAttribute("open");
  }
};

const handleEmailFormSubmit = (event) => {
  event.preventDefault();
  const form = elements.emailForm;
  if (!form) return;
  const titleField = form.elements.title;
  const title = titleField?.value?.trim() ?? "";
  if (!title) {
    if (elements.emailError) {
      elements.emailError.textContent = "Title is required.";
    }
    titleField?.focus();
    return;
  }

  const projectId = form.elements.projectId.value || "inbox";
  ensureSectionForProject(projectId);
  const sectionId = getDefaultSectionId(projectId);
  const links = collectLinks(elements.emailLinksList);
  const existing = state.tasks.find((task) => task.id === state.editingEmailId);
  const metadata = existing?.metadata && typeof existing.metadata === "object" ? { ...existing.metadata } : {};
  metadata.emailAddress = form.elements.emailAddress.value.trim();
  metadata.status = form.elements.status.value;
  metadata.links = links;

  const payload = {
    title,
    description: form.elements.notes?.value?.trim() ?? "",
    dueDate: form.elements.date.value || "",
    priority: form.elements.priority.value || "medium",
    projectId,
    sectionId,
    departmentId: form.elements.department.value || "",
    kind: "email",
    source: existing?.source ?? "manual",
    metadata,
    links,
  };

  if (state.editingEmailId && existing) {
    updateTask(state.editingEmailId, payload);
  } else {
    addTask(payload);
  }

  closeEmailDialog();
};

const handleEmailFormClick = (event) => {
  const action = event.target.dataset.action;
  if (action === "email-add-link") {
    event.preventDefault();
    createLinkRow(elements.emailLinksList);
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
  if (action === "close-email") {
    event.preventDefault();
    closeEmailDialog();
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
    });

    resetQuickAddForm();
    closeQuickAddForm();
  } catch (error) {
    console.error("Failed to create task.", error);
    elements.quickAddError.textContent = "Unable to create task. Please try again.";
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
  syncQuickAddSelectors();
  const descriptionField = elements.quickAddForm?.elements?.description;
  if (descriptionField) {
    autoResizeTextarea(descriptionField);
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
  autoResizeTextarea(elements.dialogForm.description);
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
  if (typeof elements.taskDialog.close === "function") {
    elements.taskDialog.close();
  } else {
    elements.taskDialog.removeAttribute("open");
  }
};

const navigateToTask = (taskId) => {
  const task = state.tasks.find((entry) => entry.id === taskId);
  if (!task) return;
  const project = getProjectById(task.projectId);
  const companyId = project?.companyId ?? DEFAULT_COMPANY.id;

  if (state.activeCompanyId !== companyId) {
    setActiveCompany(companyId);
  }

  if (task.projectId === "inbox") {
    setActiveView("view", "workspace");
  } else {
    setActiveView("project", task.projectId);
  }

  state.activeTaskTab = "active";
  state.activeAllTab = "created";

  if (state.viewMode !== "list") {
    setViewMode("list");
  }

  window.requestAnimationFrame(() => openTaskDialog(taskId));
};

const handleDialogSubmit = (event) => {
  event.preventDefault();
  if (!state.editingTaskId) return;
  const data = new FormData(elements.dialogForm);
  const title = normaliseTitle(data.get("title") ?? "");
  if (!title) {
    window.alert("Title is required.");
    return;
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
      completed: elements.dialogForm.completed.checked,
    });
    closeTaskDialog();
  } catch (error) {
    console.error("Failed to update task.", error);
    window.alert("Unable to save changes. Please try again.");
  }
};

const handleDialogClick = (event) => {
  const action = event.target.dataset.action;
  if (action === "close") {
    closeTaskDialog();
  } else if (action === "delete" && state.editingTaskId) {
    const confirmed = window.confirm("Delete this task? This cannot be undone.");
    if (confirmed) {
      removeTask(state.editingTaskId);
      closeTaskDialog();
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
  const project = getProjectById(projectId);
  if (!project) return;
  const nextName = window.prompt("Rename project", project.name);
  if (nextName && nextName.trim() && nextName.trim() !== project.name) {
    renameProject(project.id, nextName);
  }
  if (project.isDefault) return;
  const confirmDelete = window.confirm(
    `Delete the project "${project.name}"? All sections and tasks inside will be removed.`
  );
  if (confirmDelete) {
    deleteProject(project.id);
  }
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

const getWhatsappDestination = () => {
  const company = state.companies.find(
    (entry) => entry.name?.trim().toLowerCase() === WHATSAPP_COMPANY_NAME.trim().toLowerCase(),
  );
  if (!company) {
    throw new Error(
      `Create a company named "${WHATSAPP_COMPANY_NAME}" so WhatsApp imports know where to store tasks.`,
    );
  }
  const project = state.projects.find(
    (entry) =>
      entry.companyId === company.id &&
      entry.name?.trim().toLowerCase() === WHATSAPP_PROJECT_NAME.trim().toLowerCase(),
  );
  if (!project) {
    throw new Error(
      `Create a project named "${WHATSAPP_PROJECT_NAME}" inside "${WHATSAPP_COMPANY_NAME}" for WhatsApp imports.`,
    );
  }
  const section = ensureSectionForProject(project.id);
  return { company, project, section };
};

const ensureWhatsappLookbackWindow = () => {
  if (Number.isNaN(WHATSAPP_LOOKBACK_DAYS) || WHATSAPP_LOOKBACK_DAYS <= 0) {
    return 30;
  }
  return WHATSAPP_LOOKBACK_DAYS;
};

const parseGeminiJson = (payload) => {
  if (Array.isArray(payload)) return payload;
  if (payload && Array.isArray(payload.items)) return payload.items;
  if (payload && Array.isArray(payload.actions)) return payload.actions;
  return [];
};

const callGeminiForActionItems = async ({
  chatName,
  transcript,
  allowedAssignees,
  model = GEMINI_MODEL,
  endpoint = "v1",
}) => {
  if (!GEMINI_API_KEY) {
    throw new Error("Set VITE_GEMINI_API_KEY to enable WhatsApp imports.");
  }
  if (!transcript) return { items: [], endpointUsed: endpoint, modelUsed: model };

  const normaliseModel = (value) => {
    if (!value || typeof value !== "string") return "";
    return value.replace(/^models\//i, "").replace(/:generateContent$/i, "").trim();
  };
  const isStructuredModel = (value) => /^gemini-2\./i.test(value);

  const allowedNames = allowedAssignees.map((entry) => entry.name).filter(Boolean);
  const instructions = [
    "You are an AI assistant analysing WhatsApp group conversations to extract actionable tasks.",
    `Chat name: ${chatName}`,
    "The transcript you receive already begins at the earliest message that must be processed (last recorded import or the 30-day window). Ignore anything before the first line; read every remaining message in order.",
    "Each transcript line has the form: [index] ISO_TIMESTAMP | sender: message",
    "",
    "Extraction rules",
    "• Only capture genuine action items: explicit requests, commitments, delegations, or clear plans that describe what must happen. Skip greetings, confirmations, vague intent, or questions unless they contain a concrete next step.",
    "• Treat names consistently. Remove @ prefixes, prefer real names whenever they appear anywhere in the chat, and only fall back to a phone number (wrapped in single quotes) when no name exists at all.",
    "• If several people are asked to do something, list them all in the assignee field. If nobody is clearly responsible, set assignee to null (do not invent one).",
    "• Convert any relative timing (“tomorrow”, “Friday”, “next week”) into an absolute date in US Eastern Time (America/New_York). If no timing is given, return null.",
    '• Assign one of these priorities exactly: "critical", "very-high", "high", "medium", "low", "optional".',
    "  – critical: emergency, immediate business impact, or language such as “ASAP / now”.",
    "  – very-high: explicitly urgent or blocking core work.",
    "  – high: specific short-term deadline (e.g., due in a few days) or clearly important follow-up.",
    "  – medium: normal follow-up with implied timing but not urgent.",
    "  – low: nice-to-have or loosely timed.",
    "  – optional: discretionary or informational items that the owner may choose to do.",
    "• Use the ISO timestamp from the triggering message verbatim for sourceTimestamp, and the speaker’s name for sourceSender.",
    "",
    `You also receive a list of allowed assignees: ${allowedNames.length ? allowedNames.join(", ") : "(none)"}. Prefer those names when they match the conversation.`,
    "",
    "Output format",
    "Return a JSON array (no code fences). Each object must contain exactly:",
    "",
    "{",
    '  "title": "Short imperative task summary",',
    '  "description": "One or two sentences capturing context, commitments, and next steps with names",',
    '  "assignee": "Responsible person name or null",',
    '  "dueDate": "YYYY-MM-DD if a deadline exists, otherwise null",',
    '  "priority": "critical|very-high|high|medium|low|optional",',
    '  "sourceTimestamp": "ISO timestamp from the message",',
    '  "sourceSender": "Name of the message author"',
    "}",
    "",
    "If you find no qualifying action items, return [].",
  ].join("\n");

  const schema = {
    type: "object",
    properties: {
      items: {
        type: "array",
        items: {
          type: "object",
          properties: {
            title: { type: "string" },
            description: { type: "string" },
            assignee: { type: ["string", "null"] },
            dueDate: { type: ["string", "null"] },
            priority: {
              type: "string",
              enum: ["critical", "very-high", "high", "medium", "low", "optional"],
            },
            sourceTimestamp: { type: "string" },
            sourceSender: { type: "string" },
          },
          required: ["title", "description", "assignee", "dueDate", "priority", "sourceTimestamp", "sourceSender"],
        },
      },
    },
    required: ["items"],
  };

  const contents = [
    {
      role: "user",
      parts: [
        { text: instructions },
        { text: `Transcript:\n${transcript}` },
      ],
    },
  ];

  const buildRequestBody = (targetModel) => {
    const body = {
      contents,
      generationConfig: {
        temperature: 0.2,
        topK: 32,
      },
    };
    if (isStructuredModel(targetModel)) {
      body.tools = [
        {
          functionDeclarations: [
            {
              name: "store_action_items",
              description: "Return the list of action items extracted from the WhatsApp transcript.",
              parameters: schema,
            },
          ],
        },
      ];
      body.toolConfig = { functionCall: { name: "store_action_items" } };
    } else {
      body.generationConfig.responseMimeType = "application/json";
    }
    return body;
  };

  const runRequest = async (targetEndpoint, targetModel) => {
    const requestBody = buildRequestBody(targetModel);
    const url = `https://generativelanguage.googleapis.com/${targetEndpoint}/models/${encodeURIComponent(targetModel)}:generateContent?key=${encodeURIComponent(GEMINI_API_KEY)}`;
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(requestBody),
    });
    const payload = await response.json();
    return { response, payload };
  };

  const isRetryableModelError = (response, payload) => {
    if (!response || response.ok) return false;
    if (response.status === 404) return true;
    const message = (payload?.error?.message || "").toLowerCase();
    return message.includes("not found") || message.includes("not supported");
  };

  const initialEndpoint = endpoint === "v1beta" ? "v1beta" : "v1";
  const endpointCandidates = initialEndpoint === "v1" ? ["v1", "v1beta"] : ["v1beta", "v1"];

  const modelCandidates = (() => {
    const candidates = [];
    const seen = new Set();
    const addCandidate = (value) => {
      const trimmed = (value || "").trim();
      if (!trimmed || seen.has(trimmed)) return;
      const cleaned = normaliseModel(trimmed);
      if (!cleaned || seen.has(cleaned)) return;
      seen.add(cleaned);
      candidates.push(cleaned);
    };
    addCandidate(model);
    if (/-latest$/i.test(model)) {
      addCandidate(model.replace(/-latest$/i, ""));
    }
    if (/-\d+$/i.test(model)) {
      addCandidate(model.replace(/-\d+$/i, ""));
    }
    if (/-\d+[a-z]+$/i.test(model)) {
      addCandidate(model.replace(/-\d+[a-z]+$/i, ""));
    }
    if (/-preview$/i.test(model) || /-exp$/i.test(model)) {
      addCandidate(model.replace(/-(preview|exp)$/i, ""));
    }
    return candidates;
  })();

  let lastError = null;

  for (const modelCandidate of modelCandidates) {
    for (const endpointCandidate of endpointCandidates) {
      if (isStructuredModel(modelCandidate) && endpointCandidate !== "v1beta") {
        continue;
      }
      const { response, payload } = await runRequest(endpointCandidate, modelCandidate);
      const message = payload?.error?.message || "Gemini API request failed.";
      if (response.ok) {
        const candidate = payload?.candidates?.[0];
        const parts = candidate?.content?.parts ?? [];
        if (isStructuredModel(modelCandidate)) {
          const functionPart = parts.find((part) => part?.functionCall);
          if (!functionPart?.functionCall) {
            console.error("Gemini response missing functionCall part", payload);
            throw new Error("Gemini returned an unexpected response.");
          }
          const args = functionPart.functionCall.args ?? {};
          let rawItems;
          if (Array.isArray(args)) {
            rawItems = args;
          } else if (Array.isArray(args?.items)) {
            rawItems = args.items;
          } else if (typeof args?.items === "string") {
            try {
              rawItems = JSON.parse(args.items);
            } catch (error) {
              console.error("Failed to parse Gemini function payload", error, args.items);
              throw new Error("Gemini returned an unexpected response.");
            }
          } else if (typeof args === "string") {
            try {
              const parsed = JSON.parse(args);
              rawItems = parsed?.items ?? parsed;
            } catch (error) {
              console.error("Failed to parse Gemini function payload", error, args);
              throw new Error("Gemini returned an unexpected response.");
            }
          } else if (args && typeof args === "object") {
            rawItems = args.items ?? args;
          }
          return {
            items: parseGeminiJson(rawItems),
            endpointUsed: endpointCandidate,
            modelUsed: modelCandidate,
          };
        }

        const partText =
          candidate?.content?.parts?.[0]?.text ??
          candidate?.content?.[0]?.text ??
          (candidate?.content?.parts ?? [])
            .map((part) => part.text)
            .filter(Boolean)
            .join("\n");
        if (!partText) {
          return {
            items: [],
            endpointUsed: endpointCandidate,
            modelUsed: modelCandidate,
          };
        }

        try {
          const parsed = JSON.parse(partText);
          return {
            items: parseGeminiJson(parsed),
            endpointUsed: endpointCandidate,
            modelUsed: modelCandidate,
          };
        } catch (error) {
          console.error("Failed to parse Gemini response", error, partText);
          throw new Error("Gemini returned an unexpected response.");
        }
      }

      if (!isRetryableModelError(response, payload)) {
        throw new Error(message);
      }

      lastError = message;
    }
  }

  throw new Error(lastError || "Gemini API request failed.");
};

const summariseImportStats = ({ chatName, messageCount, taskCount, model, endpoint }) => {
  const lines = [];
  lines.push(`${messageCount} new message${messageCount === 1 ? "" : "s"} analysed`);
  lines.push(`${taskCount} action item${taskCount === 1 ? "" : "s"} created`);
  lines.push(`Source chat: ${chatName}`);
  if (model) {
    const suffix = endpoint ? ` (${endpoint})` : "";
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

const resetWhatsappImport = () => {
  const model = state.importJob.model || GEMINI_MODEL;
  const endpoint = state.importJob.endpoint || WHATSAPP_DEFAULT_ENDPOINT;
  state.importJob = {
    file: null,
    status: "idle",
    error: "",
    stats: null,
    model,
    endpoint,
  };
  if (elements.whatsappForm) {
    elements.whatsappForm.reset();
  }
  if (elements.whatsappModel) {
    elements.whatsappModel.value = model;
  }
  if (elements.whatsappEndpoint) {
    elements.whatsappEndpoint.value = endpoint;
  }
  renderWhatsappImport();
};

const renderWhatsappImport = () => {
  const { file, status, error, stats, model, endpoint } = state.importJob;
  if (elements.whatsappPreview) {
    if (stats && file) {
      elements.whatsappPreview.hidden = false;
      if (elements.whatsappFileLabel) {
        elements.whatsappFileLabel.textContent = file.name;
      }
      if (elements.whatsappRangeLabel) {
        elements.whatsappRangeLabel.textContent = stats.range ?? "Ready to analyse";
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
  if (elements.whatsappEndpoint) {
    elements.whatsappEndpoint.value = endpoint;
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
  resetWhatsappImport();
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
  const selectedModel = (state.importJob.model || GEMINI_MODEL).trim() || GEMINI_MODEL;
  const selectedEndpointRaw = (state.importJob.endpoint || WHATSAPP_DEFAULT_ENDPOINT).toLowerCase();
  const selectedEndpoint = selectedEndpointRaw === "v1beta" ? "v1beta" : "v1";
  const { chatName, messages } = await readWhatsappExport(file);
  if (!messages.length) {
    throw new Error("No messages were found in the chat export.");
  }

  const lookbackDays = ensureWhatsappLookbackWindow();
  const windowStart = new Date(Date.now() - lookbackDays * MS_IN_DAY);
  const chatKey = chatName || file.name || "default-chat";
  const lastProcessedISO = state.imports.whatsapp[chatKey];
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
    const summary = [
      lastProcessedISO
        ? `No new messages since ${new Date(lastProcessedISO).toLocaleString()}`
        : "No recent messages found in the last 30 days"
    ];
    const endpointLabel = selectedEndpoint ? " (" + selectedEndpoint + ")" : "";
    summary.push(`Model: ${selectedModel}${endpointLabel}`);
    state.importJob.model = selectedModel;
    state.importJob.endpoint = selectedEndpoint;
    state.importJob.stats = {
      range: formatDateRange(windowStart, new Date()),
      summary,
    };
    renderWhatsappImport();
    return;
  }

  const limited = filtered.slice(-Math.max(10, Math.min(filtered.length, MAX_WHATSAPP_LINES || 2000)));
  const transcript = buildTranscript(limited);
  const allowedAssignees = state.members.map((member) => ({ id: member.id, name: member.name }));
  const { items: actionItems, endpointUsed, modelUsed } = await callGeminiForActionItems({
    chatName,
    transcript,
    allowedAssignees,
    model: selectedModel,
    endpoint: selectedEndpoint,
  });
  const effectiveEndpoint = endpointUsed || selectedEndpoint;
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

    const assigneeName = typeof item.assignee === "string" ? item.assignee.trim().toLowerCase() : "";
    const assigneeId = assigneeName && memberLookup.has(assigneeName) ? memberLookup.get(assigneeName) : "";

    const dueDate = normaliseDueDate(item.dueDate);
    const priority = normalisePriority(item.priority);
    const sender = typeof item.sourceSender === "string" ? item.sourceSender.trim() : "";
    const detail = typeof item.description === "string" ? item.description.trim() : "";

    const messageSummary = sender
      ? `Source: ${sender} in ${chatName}`
      : `Source chat: ${chatName}`;
    const descriptionSegments = [detail, messageSummary];
    if (sourceTimestamp && !Number.isNaN(sourceTimestamp.getTime())) {
      descriptionSegments.push(`Mentioned on ${sourceTimestamp.toLocaleString()}`);
    }

    const task = addTask({
      title,
      description: descriptionSegments.filter(Boolean).join("\n\n"),
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
  saveImports();

  state.importJob.model = effectiveModel;
  state.importJob.endpoint = effectiveEndpoint;
  state.importJob.stats = {
    range: formatDateRange(earliestTimestamp, latestTimestamp),
    summary: summariseImportStats({
      chatName,
      messageCount: filtered.length,
      taskCount: createdTasks.length,
      model: effectiveModel,
      endpoint: effectiveEndpoint,
    }),
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
  state.importJob.model = value || GEMINI_MODEL;
  renderWhatsappImport();
};

const handleWhatsappEndpointChange = (event) => {
  const value = (event.target?.value || "").toLowerCase();
  state.importJob.endpoint = value === "v1beta" ? "v1beta" : "v1";
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
  elements.companyDialog?.addEventListener("cancel", (event) => {
    event.preventDefault();
    closeCompanyDialog();
  });
  elements.meetingForm?.addEventListener("submit", handleMeetingFormSubmit);
  elements.meetingForm?.addEventListener("click", handleMeetingFormClick);
  elements.meetingDialog?.addEventListener("cancel", (event) => {
    event.preventDefault();
    closeMeetingDialog();
  });
  elements.emailForm?.addEventListener("submit", handleEmailFormSubmit);
  elements.emailForm?.addEventListener("click", handleEmailFormClick);
  elements.emailDialog?.addEventListener("cancel", (event) => {
    event.preventDefault();
    closeEmailDialog();
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
  elements.whatsappEndpoint?.addEventListener("change", handleWhatsappEndpointChange);
  elements.whatsappForm?.addEventListener("click", handleWhatsappDialogClick);
  elements.whatsappFile?.addEventListener("change", handleWhatsappFileChange);
  elements.whatsappDialog?.addEventListener("cancel", (event) => {
    event.preventDefault();
    closeWhatsappDialog();
  });

  elements.taskTabPanels?.addEventListener("change", handleTaskCheckboxChange);
  elements.taskTabPanels?.addEventListener("click", handleTaskActionClick);

  if (elements.quickAddForm) {
    elements.quickAddForm.addEventListener("submit", handleQuickAddSubmit);
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
    closeTaskDialog();
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
  initialiseTextareaAutosize();
  await hydrateState();
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























