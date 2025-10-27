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
const MAX_WHATSAPP_LINES = Number.parseInt(import.meta.env.VITE_WHATSAPP_MAX_LINES ?? "2000", 10);
const WHATSAPP_LOG_SHEET_ID = import.meta.env.VITE_WHATSAPP_LOG_SHEET_ID || "";


const DEFAULT_COMPANY = {
  id: "company-default",
  name: "Synergy Grid",
  isDefault: true,
  createdAt: new Date().toISOString(),
};

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
  "Start by choosing a company from the dropdown under the search bar. Each company keeps a separate set of projects, sections, tasks, team metrics, and preferences.",
  "Need a new company? Open the Company dropdown, hit New company, optionally add its first project, and use the pencil/trash buttons beside any company to rename or remove it when the work wraps up.",
  "Open the Project dropdown to jump directly into work for the current company. Add projects with the inline + button, rename them with the pencil icon, or delete unused ones (trash) to keep the list lean.",
  "Stay organised with sections: click Add section while viewing a project, rename sections from their column menu, delete extras when finished, and drag section headers to reorder the workflow.",
  "Capture work through Quick Add or while editing an existing card. Pick the correct company/project/section, add due dates, priorities, departments, assignees, and attachments, then drag cards between sections as work progresses.",
  "Switch between List and Board views using the toggle in the hero card. Board view is project-specific and lets you drag cards between sections; List view works everywhere for quick triage.",
  "Moving work across companies? Edit the task (or project), choose the new project from the dropdown, and the linked company updates automatically while keeping task history intact.",
  "Use the Manage Members and Manage Departments buttons to keep team data current so you can assign tasks accurately and keep dashboard metrics meaningful.",
  "Before changing processes, export the workspace (Download button in the header) or copy the latest build so you always have a reference snapshot.",
  "Keep this Userguide current: click Userguide → Edit guide whenever instructions change so the whole team sees the latest way of working.",
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
  "Need to move work between companies? Edit the task or project and pick the new project—its company automatically follows, and tasks keep their history.",
  "Keep everyone aligned: update this Userguide whenever the process changes (Userguide → Edit guide) and export tasks regularly if you need an offline backup.",
];

const state = {
  tasks: [],
  projects: [],
  sections: [],
  companies: [],
  members: [],
  departments: [],
  settings: null,
  activeView: { type: "view", value: "inbox" },
  activeCompanyId: DEFAULT_COMPANY.id,
  companyRecents: {},
  userguide: [...DEFAULT_USERGUIDE],
  viewMode: "list",
  searchTerm: "",
  showCompleted: false,
  metricsFilter: "all",
  editingTaskId: null,
  dragTaskId: null,
  dragSectionId: null,
  sectionDropTarget: null,
  isQuickAddOpen: false,
  isUserguideOpen: false,
  isEditingUserguide: false,
  openDropdown: null,
  dialogAttachmentDraft: [],
  openSectionMenu: null,
  importJob: {
    file: null,
    status: "idle",
    error: "",
    stats: null,
  },
  imports: {
    whatsapp: {},
  },
};

const elements = {
  viewList: document.getElementById("viewList"),
  viewCounts: {
    inbox: document.querySelector('[data-count="inbox"]'),
    today: document.querySelector('[data-count="today"]'),
    upcoming: document.querySelector('[data-count="upcoming"]'),
  },
  searchInput: document.getElementById("searchInput"),
  searchInputMobile: document.getElementById("searchInput-mobile"),
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
  taskList: document.getElementById("taskList"),
  completedList: document.getElementById("completedList"),
  toggleCompleted: document.getElementById("toggleCompleted"),
  completedPlaceholder: document.querySelector('[data-empty-completed]'),
  emptyState: document.getElementById("emptyState"),
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
  companyDropdownToggle: document.getElementById("companyDropdownToggle"),
  companyDropdownMenu: document.getElementById("companyDropdownMenu"),
  companyDropdownList: document.getElementById("companyDropdownList"),
  companyDropdownLabel: document.getElementById("companyDropdownLabel"),
  projectDropdownToggle: document.getElementById("projectDropdownToggle"),
  projectDropdownMenu: document.getElementById("projectDropdownMenu"),
  projectDropdownList: document.getElementById("projectDropdownList"),
  projectDropdownLabel: document.getElementById("projectDropdownLabel"),
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
    showCompleted: state.showCompleted,
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
  if (typeof prefs.showCompleted === "boolean") {
    state.showCompleted = prefs.showCompleted;
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
const tasksForCompany = () =>
  state.tasks.filter((task) => {
    if (!state.activeCompanyId) return true;
    if (!task.companyId) return state.activeCompanyId === DEFAULT_COMPANY.id;
    return task.companyId === state.activeCompanyId;
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
  const company = getCompanyById(companyId);
  if (!company || state.activeCompanyId === companyId) return;
  state.activeCompanyId = companyId;
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
    setActiveView("view", "inbox");
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
  if (!Array.isArray(task.attachments)) {
    task.attachments = [];
  }
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
    task.companyId === state.activeCompanyId ||
    (!task.companyId && state.activeCompanyId === DEFAULT_COMPANY.id);
  if (!matchesCompany) return false;

  if (type === "view") {
    if (value === "inbox") return task.projectId === "inbox";
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

  return (
    task.title.toLowerCase().includes(needle) ||
    (task.description ?? "").toLowerCase().includes(needle) ||
    (section?.name ?? "").toLowerCase().includes(needle) ||
    (member?.name ?? "").toLowerCase().includes(needle) ||
    (department?.name ?? "").toLowerCase().includes(needle) ||
    (project?.name ?? "").toLowerCase().includes(needle) ||
    attachments.some((attachment) => (attachment.name ?? "").toLowerCase().includes(needle))
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
  renderTasks();
};

const tasksForCurrentView = () =>
  state.tasks.filter((task) => matchesActiveView(task) && matchesSearch(task));

const openTasks = (tasks) => tasks.filter((task) => !task.completed);
const completedTasks = (tasks) => tasks.filter((task) => task.completed);

const renderCompanyDropdown = () => {
  if (!elements.companyDropdownList) return;
  const fragment = document.createDocumentFragment();
  if (!state.companies.length) {
    const empty = document.createElement("li");
    empty.className = "text-sm text-slate-500 px-2 py-1.5";
    empty.textContent = "Add your first company.";
    fragment.append(empty);
  } else {
    state.companies.forEach((company) => {
      const item = document.createElement("li");
      item.className = "selector-item";
      const select = document.createElement("button");
      select.type = "button";
      select.dataset.select = "company";
      select.dataset.companyId = company.id;
      select.setAttribute("role", "option");
      const isActive = company.id === state.activeCompanyId;
      select.classList.toggle("active", isActive);
      select.setAttribute("aria-selected", String(isActive));
      select.textContent = company.name;
      const meta = document.createElement("small");
      const projectCount = getProjectsForCompany(company.id).length;
      meta.textContent = projectCount
        ? `${projectCount} project${projectCount === 1 ? "" : "s"}`
        : "No projects yet";
      select.append(meta);

      const actions = document.createElement("div");
      actions.className = "selector-item-actions";
      const edit = document.createElement("button");
      edit.type = "button";
      edit.className = "selector-action-btn";
      edit.dataset.action = "edit-company";
      edit.dataset.companyId = company.id;
      edit.setAttribute("aria-label", `Rename ${company.name}`);
      edit.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><path d="M4 20h4l10-10-4-4L4 16v4Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><path d="m14 6 4 4" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/></svg>';
      const remove = document.createElement("button");
      remove.type = "button";
      remove.className = "selector-action-btn";
      remove.dataset.action = "delete-company";
      remove.dataset.companyId = company.id;
      remove.setAttribute("aria-label", `Delete ${company.name}`);
      remove.disabled = Boolean(company.isDefault);
      remove.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><path d="M6 7h12m-9 0V5h6v2m-1 3v7m-4-7v7M7 7l1 12h8l1-12" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>';
      actions.append(edit, remove);
      item.append(select, actions);
      fragment.append(item);
    });
  }
  elements.companyDropdownList.replaceChildren(fragment);
  if (elements.companyDropdownLabel) {
    const activeCompany = getCompanyById(state.activeCompanyId);
    elements.companyDropdownLabel.textContent = activeCompany ? activeCompany.name : "Select company";
  }
};

const renderProjectDropdown = () => {
  if (!elements.projectDropdownList) return;
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
      const edit = document.createElement("button");
      edit.type = "button";
      edit.className = "selector-action-btn";
      edit.dataset.action = "edit-project";
      edit.dataset.projectId = project.id;
      edit.setAttribute("aria-label", `Rename ${project.name}`);
      edit.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><path d="M4 20h4l10-10-4-4L4 16v4Z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><path d="m14 6 4 4" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/></svg>';
      const remove = document.createElement("button");
      remove.type = "button";
      remove.className = "selector-action-btn";
      remove.dataset.action = "delete-project";
      remove.dataset.projectId = project.id;
      remove.setAttribute("aria-label", `Delete ${project.name}`);
      remove.disabled = Boolean(project.isDefault);
      remove.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none"><path d="M6 7h12m-9 0V5h6v2m-1 3v7m-4-7v7M7 7l1 12h8l1-12" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>';
      actions.append(edit, remove);
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
    inbox: scopedTasks.filter((task) => task.projectId === "inbox" && !task.completed).length,
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

  elements.viewCounts.inbox.textContent = counts.inbox;
  elements.viewCounts.today.textContent = counts.today;
  elements.viewCounts.upcoming.textContent = counts.upcoming;
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
    option.textContent = company && !company.isDefault ? `${project.name} · ${company.name}` : project.name;
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
  elements.viewList
    .querySelectorAll(".nav-item[data-view]")
    .forEach((button) => {
      const isActive =
        state.activeView.type === "view" && button.dataset.view === state.activeView.value;
      button.classList.toggle("active", isActive);
    });
};

const describeView = () => {
  const { type, value } = state.activeView;
  if (type === "view") {
    if (value === "inbox") {
      return { title: "Inbox", subtitle: "All unscheduled tasks live here." };
    }
    if (value === "today") {
      const formatted = new Intl.DateTimeFormat(undefined, {
        weekday: "long",
        month: "short",
        day: "numeric",
      }).format(new Date());
      return { title: "Today", subtitle: formatted };
    }
    if (value === "upcoming") {
      return { title: "Upcoming", subtitle: "Schedule ahead for the next few weeks." };
    }
  }
  if (type === "project") {
    const project = getProjectById(value);
    if (project) {
      const company = getCompanyById(project.companyId);
      return {
        title: project.name,
        subtitle: company ? `${company.name} · Project board` : "View tasks scoped to this project.",
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

  buildMeta(task, metaEl);

  if (Array.isArray(task.attachments) && task.attachments.length && contentEl) {
    const row = document.createElement("div");
    row.className = "attachments-row";
    task.attachments.forEach((attachment) => {
      row.append(createAttachmentChip(attachment));
    });
    contentEl.append(row);
  }

  return item;
};

const renderListView = (tasks) => {
  const active = openTasks(tasks).sort(compareTasks);
  const completed = completedTasks(tasks).sort(compareTasks);

  elements.taskList.replaceChildren(...active.map(renderTaskItem));
  elements.completedList.replaceChildren(...completed.map(renderTaskItem));

  elements.emptyState.hidden = active.length !== 0;
  elements.completedList.classList.toggle("hidden", !state.showCompleted);
  if (elements.completedPlaceholder) {
    elements.completedPlaceholder.hidden = completed.length !== 0;
  }
  elements.toggleCompleted.textContent = state.showCompleted ? "Hide completed" : "Show completed";
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

  const menuWrapper = document.createElement("div");
  menuWrapper.className = "section-menu-wrapper";
  const menuButton = document.createElement("button");
  menuButton.type = "button";
  menuButton.className = "section-menu-button";
  menuButton.dataset.action = "section-menu";
  menuButton.dataset.sectionId = section.id;
  menuButton.innerHTML = '<span class="sr-only">Section actions</span><svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 12h.01M12 12h.01M19 12h.01"/></svg>';
  const menu = document.createElement("div");
  menu.className = "section-menu";
  if (!section.locked && getSectionsForProject(section.projectId).length > 1) {
    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.dataset.action = "section-delete";
    deleteBtn.dataset.sectionId = section.id;
    deleteBtn.className = "danger";
    deleteBtn.textContent = "Delete section";
    menu.append(deleteBtn);
  }
  if (menu.childElementCount) {
    menuWrapper.append(menuButton, menu);
    header.append(menuWrapper);
  }

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
  closeSectionMenu();
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
  const recent = [...tasksForCompany()]
    .sort((a, b) => {
      const bTime = new Date(b.updatedAt || b.createdAt || 0).getTime();
      const aTime = new Date(a.updatedAt || a.createdAt || 0).getTime();
      return bTime - aTime;
    })
    .slice(0, 6);

  if (!recent.length) {
    const empty = document.createElement("p");
    empty.className = "text-sm text-slate-500";
    empty.textContent = "No activity yet. Add your first task.";
    fragment.append(empty);
  } else {
    recent.forEach((task) => {
      const item = document.createElement("div");
      item.className = "activity-item bg-slate-50 border border-slate-200 rounded-xl px-3 py-2";

      const title = document.createElement("p");
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

  elements.activityFeed.replaceChildren(fragment);
};

const renderUserguidePanel = () => {
  if (!elements.userguidePanel) return;
  if (elements.userguideList) {
    const fragment = document.createDocumentFragment();
    state.userguide.forEach((entry) => {
      const item = document.createElement("li");
      item.textContent = entry;
      fragment.append(item);
    });
    elements.userguideList.replaceChildren(fragment);
    elements.userguideList.hidden = state.isEditingUserguide;
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
  renderCompanyDropdown();
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
      nextValue = "inbox";
    } else {
      rememberProjectSelection(project.id);
    }
  }

  state.activeView = { type: nextType, value: nextValue };
  if (state.viewMode === "board" && nextType !== "project") {
    state.viewMode = "list";
  }
  savePreferences();
  render();
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
    completed: false,
    createdAt,
    updatedAt: now,
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
  state.tasks = state.tasks.filter((task) => task.id !== taskId);
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

  state.tasks = state.tasks.filter((task) => task.projectId !== projectId);
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
      setActiveView("view", "inbox");
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
  renderCompanyDropdown();
  renderProjectDropdown();
  return company;
};

const renameCompany = (companyId, nextName) => {
  const company = getCompanyById(companyId);
  if (!company) return;
  const trimmed = nextName.trim();
  if (!trimmed) return;
  const exists = state.companies.some(
    (entry) => entry.id !== companyId && entry.name.toLowerCase() === trimmed.toLowerCase(),
  );
  if (exists) {
    window.alert("Another company already uses that name.");
    return;
  }
  company.name = trimmed;
  company.updatedAt = new Date().toISOString();
  saveCompanies();
  renderCompanyDropdown();
  renderProjectDropdown();
  renderHeader();
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

  state.tasks = state.tasks.filter((task) => !projectIds.includes(task.projectId));
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
      setActiveView("view", "inbox");
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
  const item = button.closest(".task-item");
  if (!item) return;
  const taskId = item.dataset.taskId;

  if (action === "edit") {
    openTaskDialog(taskId);
  } else if (action === "delete") {
    const confirmed = window.confirm("Delete this task? This cannot be undone.");
    if (confirmed) {
      removeTask(taskId);
    }
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

const handleToggleCompleted = () => {
  state.showCompleted = !state.showCompleted;
  savePreferences();
  renderTasks();
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
  setActiveView("view", "inbox");
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

const handleCompanyMenuClick = (event) => {
  const addButton = event.target.closest('[data-action="add-company"]');
  if (addButton) {
    const name = window.prompt("Company name");
    if (name) {
      const company = createCompany(name);
      if (company) {
        setActiveCompany(company.id);
        const projectName = window.prompt("Add a first project for this company?");
        if (projectName) {
          createProject(projectName, company.id);
        } else {
          renderProjectDropdown();
        }
      }
    }
    closeDropdown("company");
    return;
  }

  const renameButton = event.target.closest('[data-action="edit-company"]');
  if (renameButton) {
    const company = getCompanyById(renameButton.dataset.companyId);
    if (!company) return;
    const nextName = window.prompt("Rename company", company.name);
    if (nextName) {
      renameCompany(company.id, nextName);
    }
    return;
  }

  const deleteButton = event.target.closest('[data-action="delete-company"]');
  if (deleteButton) {
    deleteCompany(deleteButton.dataset.companyId);
    closeDropdown("company");
    return;
  }

  const selectButton = event.target.closest('button[data-select="company"]');
  if (selectButton) {
    setActiveCompany(selectButton.dataset.companyId);
    closeDropdown("company");
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

  const renameButton = event.target.closest('[data-action="edit-project"]');
  if (renameButton) {
    const project = getProjectById(renameButton.dataset.projectId);
    if (!project) return;
    const nextName = window.prompt("Rename project", project.name);
    if (nextName) {
      renameProject(project.id, nextName);
    }
    return;
  }

  const deleteButton = event.target.closest('[data-action="delete-project"]');
  if (deleteButton) {
    deleteProject(deleteButton.dataset.projectId);
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
  const menuTrigger = event.target.closest('[data-action="section-menu"]');
  if (menuTrigger) {
    toggleSectionMenu(menuTrigger.dataset.sectionId, menuTrigger);
    return;
  }

  const deleteButton = event.target.closest('[data-action="section-delete"]');
  if (deleteButton) {
    closeSectionMenu();
    const sectionId = deleteButton.dataset.sectionId;
    if (sectionId) {
      deleteSection(sectionId);
    }
  }
};

const toggleSectionMenu = (sectionId, trigger) => {
  if (!elements.boardColumns) return;
  const wrapper = trigger?.closest('.section-menu-wrapper');
  const menu = wrapper?.querySelector('.section-menu');
  if (!menu) return;

  if (state.openSectionMenu && state.openSectionMenu.menu !== menu) {
    closeSectionMenu();
  }

  const isOpen = menu.classList.contains('open');
  if (isOpen) {
    closeSectionMenu();
    return;
  }

  menu.classList.add('open');
  state.openSectionMenu = { sectionId, menu };
};

const closeSectionMenu = () => {
  if (!state.openSectionMenu) return;
  state.openSectionMenu.menu.classList.remove('open');
  state.openSectionMenu = null;
};

const dropdownElements = {
  company: () => ({
    toggle: elements.companyDropdownToggle,
    menu: elements.companyDropdownMenu,
  }),
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
  return `${formatter.format(start)} – ${formatter.format(end)}`;
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

const callGeminiForActionItems = async ({ chatName, transcript, allowedAssignees }) => {
  if (!GEMINI_API_KEY) {
    throw new Error("Set VITE_GEMINI_API_KEY to enable WhatsApp imports.");
  }
  if (!transcript) return [];

  const allowedNames = allowedAssignees.map((entry) => entry.name).filter(Boolean);
  const instructions = [
    "You are extracting actionable tasks from a WhatsApp chat transcript.",
    `Chat name: ${chatName}`,
    "The transcript contains lines formatted as: [index] ISO_TIMESTAMP | sender: message",
    "Return ONLY JSON (do not wrap in code fences). The JSON must be an array of objects with these fields:",
    '- "title": short imperative summary of the action item.',
    '- "description": fuller explanation including relevant context and next steps.',
    '- "assignee": exactly match one of the allowed names below; if nobody applies, use null.',
    '- "dueDate": ISO date string (YYYY-MM-DD) if an explicit or strongly implied deadline exists, otherwise null.',
    '- "priority": one of ["critical","very-high","high","medium","low","optional"] (default to "medium").',
    '- "sourceTimestamp": copy the ISO timestamp from the transcript line that triggered the action.',
    '- "sourceSender": the name of the person who stated the action item.',
    "Only include genuine action items, commitments, or requests that require follow-up.",
    `Allowed assignees: ${allowedNames.length ? allowedNames.join(", ") : "(none)"}.`,
    "If no action items are present, return an empty JSON array.",
  ].join("\n");

  const body = {
    contents: [
      {
        role: "user",
        parts: [
          { text: instructions },
          { text: `Transcript:\n${transcript}` },
        ],
      },
    ],
    generationConfig: {
      temperature: 0.2,
      topK: 32,
      responseMimeType: "application/json",
    },
  };

  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(
    GEMINI_MODEL,
  )}:generateContent?key=${encodeURIComponent(GEMINI_API_KEY)}`;
  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  const payload = await response.json();
  if (!response.ok) {
    const message = payload?.error?.message || "Gemini API request failed.";
    throw new Error(message);
  }

  const candidate = payload?.candidates?.[0];
  const partText =
    candidate?.content?.parts?.[0]?.text ??
    candidate?.content?.[0]?.text ??
    candidate?.content?.parts?.map((part) => part.text).filter(Boolean).join("\n");
  if (!partText) return [];

  try {
    const parsed = JSON.parse(partText);
    return parseGeminiJson(parsed);
  } catch (error) {
    console.error("Failed to parse Gemini response", error, partText);
    throw new Error("Gemini returned an unexpected response.");
  }
};

const summariseImportStats = ({ chatName, messageCount, taskCount }) => {
  const lines = [];
  lines.push(`${messageCount} new message${messageCount === 1 ? "" : "s"} analysed`);
  lines.push(`${taskCount} action item${taskCount === 1 ? "" : "s"} created`);
  lines.push(`Source chat: ${chatName}`);
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
  state.importJob = {
    file: null,
    status: "idle",
    error: "",
    stats: null,
  };
  if (elements.whatsappForm) {
    elements.whatsappForm.reset();
  }
  renderWhatsappImport();
};

const renderWhatsappImport = () => {
  const { file, status, error, stats } = state.importJob;
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

  if (elements.whatsappFile) {
    elements.whatsappFile.disabled = status === "processing";
  }
  const submitButton = elements.whatsappForm?.querySelector('[data-action="submit-whatsapp"]');
  if (submitButton) {
    submitButton.disabled = status === "processing" || !file;
    submitButton.textContent =
      status === "processing" ? "Processing…" : status === "completed" ? "Run again" : "Fetch action items";
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
    state.importJob.stats = {
      range: formatDateRange(windowStart, new Date()),
      summary: [
        lastProcessedISO
          ? `No new messages since ${new Date(lastProcessedISO).toLocaleString()}`
          : "No recent messages found in the last 30 days",
      ],
    };
    renderWhatsappImport();
    return;
  }

  const limited = filtered.slice(-Math.max(10, Math.min(filtered.length, MAX_WHATSAPP_LINES || 2000)));
  const transcript = buildTranscript(limited);
  const allowedAssignees = state.members.map((member) => ({ id: member.id, name: member.name }));
  const actionItems = await callGeminiForActionItems({
    chatName,
    transcript,
    allowedAssignees,
  });

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

  state.importJob.stats = {
    range: formatDateRange(earliestTimestamp, latestTimestamp),
    summary: summariseImportStats({
      chatName,
      messageCount: filtered.length,
      taskCount: createdTasks.length,
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

const handleWhatsappDialogClick = (event) => {
  const action = event.target?.dataset?.action;
  if (action === "cancel-whatsapp") {
    event.preventDefault();
    closeWhatsappDialog();
  }
};

const handleGlobalClick = (event) => {
  if (event.target.closest('[data-action="section-menu"]')) return;
  if (event.target.closest('.section-menu')) return;
  closeSectionMenu();
  if (!event.target.closest(".selector-dropdown")) {
    closeAllDropdowns();
  }
};

const handleGlobalKeydown = (event) => {
  if (event.key === 'Escape') {
    closeSectionMenu();
    closeAllDropdowns();
    closeUserguide();
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
  elements.companyDropdownToggle?.addEventListener("click", () => toggleDropdown("company"));
  elements.projectDropdownToggle?.addEventListener("click", () => toggleDropdown("project"));
  elements.companyDropdownMenu?.addEventListener("click", handleCompanyMenuClick);
  elements.projectDropdownMenu?.addEventListener("click", handleProjectMenuClick);
  elements.userguideEditToggle?.addEventListener("click", handleUserguideEditToggle);
  elements.userguideCancelEdit?.addEventListener("click", handleUserguideCancelEdit);
  elements.userguideForm?.addEventListener("submit", handleUserguideSave);
  document
    .querySelectorAll('[data-action="open-userguide"]')
    .forEach((button) => button.addEventListener("click", toggleUserguide));
  document
    .querySelectorAll('[data-action="close-userguide"]')
    .forEach((button) => button.addEventListener("click", closeUserguide));
  elements.importWhatsapp?.addEventListener("click", openWhatsappDialog);
  elements.whatsappForm?.addEventListener("submit", handleWhatsappSubmit);
  elements.whatsappForm?.addEventListener("click", handleWhatsappDialogClick);
  elements.whatsappFile?.addEventListener("change", handleWhatsappFileChange);
  elements.whatsappDialog?.addEventListener("cancel", (event) => {
    event.preventDefault();
    closeWhatsappDialog();
  });

  [elements.taskList, elements.completedList].forEach((list) => {
    if (!list) return;
    list.addEventListener("change", handleTaskCheckboxChange);
    list.addEventListener("click", handleTaskActionClick);
  });

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
  elements.toggleCompleted?.addEventListener("click", handleToggleCompleted);
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
const LEGACY_USERGUIDES = [LEGACY_USERGUIDE_V1, LEGACY_USERGUIDE_V2];
