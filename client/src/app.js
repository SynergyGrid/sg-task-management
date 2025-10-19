const STORAGE_KEYS = {
  tasks: "synergygrid.todoist.tasks.v2",
  projects: "synergygrid.todoist.projects.v2",
  sections: "synergygrid.todoist.sections.v1",
  members: "synergygrid.todoist.members.v1",
  departments: "synergygrid.todoist.departments.v1",
  preferences: "synergygrid.todoist.preferences.v2",
  settings: "synergygrid.todoist.settings.v1",
};


const DEFAULT_PROJECT = {
  id: "inbox",
  name: "Inbox",
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

const state = {
  tasks: [],
  projects: [],
  sections: [],
  members: [],
  departments: [],
  settings: null,
  activeView: { type: "view", value: "inbox" },
  viewMode: "list",
  searchTerm: "",
  showCompleted: false,
  metricsFilter: "all",
  editingTaskId: null,
  dragTaskId: null,
  dragSectionId: null,
  sectionDropTarget: null,
  isQuickAddOpen: false,
  dialogAttachmentDraft: [],
  openSectionMenu: null,
};


const elements = {
  viewList: document.getElementById("viewList"),
  projectList: document.getElementById("projectList"),
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
  projectTemplate: document.getElementById("projectItemTemplate"),
  activeTasksMetric: document.getElementById("active-tasks"),
  activityFeed: document.getElementById("activity-feed"),
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

const saveTasks = () => saveJSON(STORAGE_KEYS.tasks, state.tasks);
const saveProjects = () => saveJSON(STORAGE_KEYS.projects, state.projects);
const saveSections = () => saveJSON(STORAGE_KEYS.sections, state.sections);
const saveMembers = () => saveJSON(STORAGE_KEYS.members, state.members);
const saveDepartments = () => saveJSON(STORAGE_KEYS.departments, state.departments);
const savePreferences = () =>
  saveJSON(STORAGE_KEYS.preferences, {
    activeView: state.activeView,
    viewMode: state.viewMode,
    showCompleted: state.showCompleted,
    metricsFilter: state.metricsFilter,
  });

const generateId = (prefix) => {
  const fallback = `${prefix}-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
  return window.crypto?.randomUUID?.() ?? fallback;
};

const pickProjectColor = (index) => PROJECT_COLORS[index % PROJECT_COLORS.length];

const getProjectById = (projectId) => state.projects.find((project) => project.id === projectId);
const getSectionById = (sectionId) => state.sections.find((section) => section.id === sectionId);
const getSectionsForProject = (projectId) =>
  state.sections
    .filter((section) => section.projectId === projectId)
    .sort((a, b) => (a.order ?? 0) - (b.order ?? 0));

const getMemberById = (memberId) => state.members.find((member) => member.id === memberId);
const getDepartmentById = (departmentId) =>
  state.departments.find((department) => department.id === departmentId);

const ensureDefaultProject = () => {
  const hasInbox = state.projects.some((project) => project.id === DEFAULT_PROJECT.id);
  if (!hasInbox) {
    state.projects.unshift({ ...DEFAULT_PROJECT });
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

const renderProjects = () => {
  elements.projectList.replaceChildren();
  const fragment = document.createDocumentFragment();
  state.projects
    .filter((project) => !project.isDefault)
    .forEach((project) => {
      const template = elements.projectTemplate.content.cloneNode(true);
      const button = template.querySelector("button");
      const dot = template.querySelector(".dot");
      const label = template.querySelector(".label");
      const count = template.querySelector(".count");

      button.dataset.project = project.id;
      label.textContent = project.name;
      dot.style.background = project.color ?? PROJECT_COLORS[0];
      count.textContent = state.tasks.filter(
        (task) => task.projectId === project.id && !task.completed
      ).length;

      if (state.activeView.type === "project" && state.activeView.value === project.id) {
        button.classList.add("active");
      }

      fragment.append(template);
    });
  elements.projectList.append(fragment);
};

const updateViewCounts = () => {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const counts = {
    inbox: state.tasks.filter((task) => task.projectId === "inbox" && !task.completed).length,
    today: state.tasks.filter((task) => {
      if (!task.dueDate || task.completed) return false;
      const due = new Date(task.dueDate);
      due.setHours(0, 0, 0, 0);
      return due.getTime() === today.getTime();
    }).length,
    upcoming: state.tasks.filter((task) => {
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
    option.textContent = project.name;
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

  elements.projectList
    .querySelectorAll(".nav-item[data-project]")
    .forEach((button) => {
      const isActive =
        state.activeView.type === "project" && button.dataset.project === state.activeView.value;
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
      return { title: project.name, subtitle: "View tasks scoped to this project." };
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
  const summaries = state.departments.map((department) => {
    const members = state.members.filter((member) => member.departmentId === department.id);
    const activeTasks = state.tasks.filter((task) => task.departmentId === department.id && !task.completed).length;
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
  const recent = [...state.tasks]
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

const updateDashboardMetrics = () => {
  const filter = state.metricsFilter || 'all';
  let value = 0;

  if (filter === 'completed') {
    value = state.tasks.filter((task) => task.completed).length;
  } else if (filter === 'all') {
    value = state.tasks.filter((task) => !task.completed).length;
  } else {
    value = state.tasks.filter((task) => !task.completed && task.priority === filter).length;
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
  renderProjects();
  updateActiveNav();
  updateViewCounts();
  renderActivityFeed();
};

const syncQuickAddSelectors = () => {
  const defaultProjectId =
    state.activeView.type === "project" ? state.activeView.value : "inbox";
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
  state.activeView = { type, value };
  if (state.viewMode === "board" && type !== "project") {
    state.viewMode = "list";
  }
  savePreferences();
  render();
};
const addTask = (payload) => {
  const projectId = payload.projectId || "inbox";
  ensureSectionForProject(projectId);

  const task = {
    id: generateId("task"),
    title: payload.title,
    description: payload.description,
    dueDate: payload.dueDate,
    priority: payload.priority,
    projectId,
    sectionId: payload.sectionId || getDefaultSectionId(projectId),
    departmentId: payload.departmentId || "",
    assigneeId: payload.assigneeId || "",
    attachments: Array.isArray(payload.attachments) ? payload.attachments : [],
    completed: false,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };

  state.tasks.push(task);
  saveTasks();
  render();
};

const updateTask = (taskId, updates) => {
  const index = state.tasks.findIndex((task) => task.id === taskId);
  if (index === -1) return;
  const previous = state.tasks[index];
  const nextProjectId = updates.projectId ?? previous.projectId;
  ensureSectionForProject(nextProjectId);
  const nextSectionId =
    updates.sectionId && getSectionById(updates.sectionId)
      ? updates.sectionId
      : previous.sectionId || getDefaultSectionId(nextProjectId);

  const nextAttachments = Array.isArray(updates.attachments)
    ? updates.attachments
    : previous.attachments || [];

  state.tasks[index] = {
    ...previous,
    ...updates,
    projectId: nextProjectId,
    sectionId: nextSectionId,
    attachments: nextAttachments,
    updatedAt: new Date().toISOString(),
  };
  saveTasks();
  render();
};

const removeTask = (taskId) => {
  state.tasks = state.tasks.filter((task) => task.id !== taskId);
  saveTasks();
  render();
};

const createProject = (name) => {
  const trimmed = name.trim();
  if (!trimmed) return;
  const exists = state.projects.some(
    (project) => project.name.toLowerCase() === trimmed.toLowerCase()
  );
  if (exists) {
    window.alert("A project with this name already exists.");
    return;
  }
  const project = {
    id: generateId("project"),
    name: trimmed,
    color: pickProjectColor(state.projects.length),
    createdAt: new Date().toISOString(),
  };
  state.projects.push(project);
  saveProjects();
  ensureSectionForProject(project.id);
  saveSections();
  setActiveView("project", project.id);
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
  if (!section) return;

  const sections = getSectionsForProject(section.projectId);
  if (sections.length <= 1) {
    window.alert("A project must have at least one section.");
    return;
  }

  const fallback = sections.find((entry) => entry.id !== sectionId)?.id;
  state.tasks = state.tasks.map((task) =>
    task.sectionId === sectionId ? { ...task, sectionId: fallback } : task
  );
  state.sections = state.sections.filter((entry) => entry.id !== sectionId);

  saveSections();
  saveTasks();
  render();
};

const addMember = (name, departmentId) => {
  const trimmed = name.trim();
  if (!trimmed) return;
  const member = {
    id: generateId("member"),
    name: trimmed,
    departmentId: departmentId || "",
    createdAt: new Date().toISOString(),
  };
  state.members.push(member);
  saveMembers();
  updateTeamSelects();
  renderMemberList();
  renderTeamStatus();
  updateDashboardMetrics();
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
    createdAt: new Date().toISOString(),
  };
  state.departments.push(department);
  saveDepartments();
  updateTeamSelects();
  renderDepartmentList();
  renderTeamStatus();
  updateDashboardMetrics();
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
  updateDashboardMetrics();
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

const handleProjectClick = (event) => {
  const button = event.target.closest("[data-project]");
  if (!button) return;
  setActiveView("project", button.dataset.project);
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
    if (confirmed) removeTask(taskId);
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
  const name = window.prompt("Project name");
  if (!name) return;
  createProject(name);
};

const handleExport = () => {
  const payload = {
    tasks: state.tasks,
    projects: state.projects,
    sections: state.sections,
    members: state.members,
    departments: state.departments,
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
  state.members = [];
  state.departments = [{ ...DEFAULT_DEPARTMENT }];

  ensureSectionForProject("inbox");

  saveTasks();
  saveProjects();
  saveSections();
  saveMembers();
  saveDepartments();
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
  createSection(state.activeView.value, name);
  saveSections();
  render();
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

const handleGlobalClick = (event) => {
  if (event.target.closest('[data-action="section-menu"]')) return;
  if (event.target.closest('.section-menu')) return;
  closeSectionMenu();
};

const handleGlobalKeydown = (event) => {
  if (event.key === 'Escape') {
    closeSectionMenu();
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
    addMember(name, departmentId);
    elements.membersForm.reset();
    elements.membersForm.memberDepartment.value = DEFAULT_DEPARTMENT.id;
  }
};

const handleMembersFormClick = (event) => {
  const action = event.target.dataset.action;
  if (action === "close") {
    closeMembersDialog();
    return;
  }
  if (action === "remove-member") {
    removeMember(event.target.dataset.memberId);
  }
};

const handleDepartmentsFormSubmit = (event) => {
  event.preventDefault();
  const action = event.submitter?.dataset.action;
  if (action === "add-department") {
    const name = elements.departmentsForm.departmentName.value;
    addDepartment(name);
    elements.departmentsForm.reset();
  }
};

const handleDepartmentsFormClick = (event) => {
  const action = event.target.dataset.action;
  if (action === "close") {
    closeDepartmentsDialog();
    return;
  }
  if (action === "remove-department") {
    removeDepartment(event.target.dataset.departmentId);
  }
};
const hydrateState = () => {
  state.tasks = loadJSON(STORAGE_KEYS.tasks, []);
  state.projects = loadJSON(STORAGE_KEYS.projects, []);
  state.sections = loadJSON(STORAGE_KEYS.sections, []);
  state.members = loadJSON(STORAGE_KEYS.members, []);
  state.departments = loadJSON(STORAGE_KEYS.departments, []);
  state.settings = normaliseSettings(loadJSON(STORAGE_KEYS.settings, defaultSettings()));

  ensureDefaultProject();
  ensureDefaultDepartment();
  ensureAllProjectsHaveSections();

  state.tasks.forEach(ensureTaskDefaults);
  saveTasks();
  saveSections();
  saveDepartments();

  const prefs = loadJSON(STORAGE_KEYS.preferences, {});
  if (prefs.activeView?.type && prefs.activeView?.value) {
    state.activeView = prefs.activeView;
  }
  if (prefs.viewMode === "board" || prefs.viewMode === "list") {
    state.viewMode = prefs.viewMode;
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

  applySettings();
};

const registerEventListeners = () => {
  elements.viewList?.addEventListener("click", handleViewClick);
  elements.projectList?.addEventListener("click", handleProjectClick);
  elements.boardColumns?.addEventListener("click", handleBoardClick);

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

const init = () => {
  hydrateState();
  registerEventListeners();
  render();
};

document.addEventListener("DOMContentLoaded", init);


