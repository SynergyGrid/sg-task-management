import { db } from './lib/firebaseClient';
import { doc, getDoc, setDoc } from 'firebase/firestore';

const STORAGE_KEYS = {
  settings: "synergygrid.todoist.settings.v1",
  members: "synergygrid.todoist.members.v1",
  departments: "synergygrid.todoist.departments.v1",
  tasks: "synergygrid.todoist.tasks.v2",
  projects: "synergygrid.todoist.projects.v2",
  sections: "synergygrid.todoist.sections.v1",
  companies: "synergygrid.todoist.companies.v1",
  userguide: "synergygrid.todoist.userguide.v1",
};

const WORKSPACE_ID = import.meta.env.VITE_FIREBASE_WORKSPACE_ID ?? 'default';
const workspaceRef = doc(db, 'workspaces', WORKSPACE_ID);

const FONT_SCALE_MIN = 0.9;
const FONT_SCALE_MAX = 1.15;
const clampFontScale = (value) => {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return 1;
  if (numeric < FONT_SCALE_MIN) return FONT_SCALE_MIN;
  if (numeric > FONT_SCALE_MAX) return FONT_SCALE_MAX;
  return numeric;
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

const normaliseSettings = (settings = {}) => {
  const defaults = defaultSettings();
  return {
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
};

const DEFAULT_DEPARTMENT = {
  id: "department-general",
  name: "General",
  isDefault: true,
  createdAt: new Date().toISOString(),
};

const generateId = (prefix) => {
  const fallback = `${prefix}-${Date.now()}-${Math.floor(Math.random() * 1000)}`;
  return window.crypto?.randomUUID?.() ?? fallback;
};

const hexToRgba = (hex, alpha) => {
  if (!hex) return `rgba(37, 99, 235, ${alpha})`;
  const normalized = (hex || "").replace("#", "");
  const extended = normalized.length === 3
    ? normalized.split("").map((char) => char + char).join("")
    : normalized;
  const bigint = Number.parseInt(extended || "000000", 16);
  const r = (bigint >> 16) & 255;
  const g = (bigint >> 8) & 255;
  const b = bigint & 255;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
};

const loadLocalSettings = () => {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEYS.settings);
    if (!raw) return defaultSettings();
    return normaliseSettings(JSON.parse(raw));
  } catch {
    return defaultSettings();
  }
};

const saveLocalSettings = (settings) => {
  try {
    window.localStorage.setItem(STORAGE_KEYS.settings, JSON.stringify(settings));
  } catch {
    // ignore localStorage errors
  }
};

const saveSettingsRemote = async (settings) => {
  try {
    await setDoc(
      workspaceRef,
      {
        settings,
        updatedAt: new Date().toISOString(),
      },
      { merge: true },
    );
  } catch (error) {
    console.error('Failed to save settings to Firestore', error);
  }
};

const readLocalJSON = (key, fallback) => {
  try {
    const raw = window.localStorage.getItem(key);
    if (!raw) return fallback;
    const parsed = JSON.parse(raw);
    return parsed ?? fallback;
  } catch {
    return fallback;
  }
};

const fetchRemoteSettings = async () => {
  try {
    const snapshot = await getDoc(workspaceRef);
    if (snapshot.exists()) {
      const remoteSettings = snapshot.data().settings;
      if (remoteSettings) {
        return normaliseSettings(remoteSettings);
      }
    } else {
      await setDoc(workspaceRef, { settings: defaultSettings() }, { merge: true });
    }
  } catch (error) {
    console.error('Failed to load settings from Firestore', error);
  }
  return null;
};

const elements = {
  form: document.getElementById("settingsForm"),
  profileName: document.getElementById("profileName"),
  profilePhoto: document.getElementById("profilePhoto"),
  accentColor: document.getElementById("accentColor"),
  priorityCritical: document.getElementById("priorityCritical"),
  priorityVeryHigh: document.getElementById("priorityVeryHigh"),
  priorityHigh: document.getElementById("priorityHigh"),
  priorityMedium: document.getElementById("priorityMedium"),
  priorityLow: document.getElementById("priorityLow"),
  priorityOptional: document.getElementById("priorityOptional"),
  profilePreviewAvatar: document.getElementById("profilePreviewAvatar"),
  profilePreviewName: document.getElementById("profilePreviewName"),
  priorityPreviewChips: document.querySelectorAll("[data-preview-chip]"),
  resetButton: document.getElementById("resetSettings"),
  saveStatus: document.getElementById("saveStatus"),
  memberList: document.getElementById("settingsMemberList"),
  memberForm: document.getElementById("settingsMemberForm"),
  memberError: document.querySelector('[data-member-error]'),
  memberDepartment: document.querySelector('#settingsMemberForm select[name="memberDepartment"]'),
  departmentList: document.getElementById("settingsDepartmentList"),
  departmentForm: document.getElementById("settingsDepartmentForm"),
  departmentError: document.querySelector('[data-department-error]'),
  exportWorkspace: document.getElementById("exportWorkspace"),
  fontScale: document.getElementById("fontScale"),
  fontScaleValue: document.getElementById("fontScaleValue"),
};

let teamMembers = [];
let teamDepartments = [];

let draft = normaliseSettings(loadLocalSettings());

const applyThemePreview = (settings) => {
  const root = document.documentElement;
  root.style.setProperty("--accent", settings.theme.accent);
  root.style.setProperty("--accent-soft", hexToRgba(settings.theme.accent, 0.15));
  root.style.setProperty("--accent-strong", hexToRgba(settings.theme.accent, 0.25));
  const { priorities } = settings.theme;
  root.style.setProperty("--priority-critical", priorities.critical);
  root.style.setProperty("--priority-very-high", priorities.veryHigh);
  root.style.setProperty("--priority-high", priorities.high);
  root.style.setProperty("--priority-medium", priorities.medium);
  root.style.setProperty("--priority-low", priorities.low);
  root.style.setProperty("--priority-optional", priorities.optional);

  elements.priorityPreviewChips.forEach((chip) => {
    const key = chip.dataset.previewChip;
    if (!key) return;
    const color = priorities[key] || settings.theme.accent;
    chip.style.setProperty("background", hexToRgba(color, 0.18));
    chip.style.setProperty("border-color", hexToRgba(color, 0.25));
    chip.style.setProperty("color", color);
  });
};

const applyProfilePreview = (settings) => {
  elements.profilePreviewName.textContent = settings.profile.name;
  elements.profilePreviewAvatar.src = settings.profile.photo;
};

const updateFontScaleControl = (scale) => {
  const percent = Math.round(scale * 100);
  if (elements.fontScale) {
    elements.fontScale.value = percent;
  }
  if (elements.fontScaleValue) {
    elements.fontScaleValue.textContent = `${percent}%`;
  }
};

const applyFontScalePreview = (settings) => {
  const scale = clampFontScale(settings.display?.fontScale || 1);
  document.documentElement.style.setProperty("--workspace-font-scale", scale.toString());
  updateFontScaleControl(scale);
};

const populateForm = (settings) => {
  elements.profileName.value = settings.profile.name;
  elements.profilePhoto.value = settings.profile.photo;
  elements.accentColor.value = settings.theme.accent;
  elements.priorityCritical.value = settings.theme.priorities.critical;
  elements.priorityVeryHigh.value = settings.theme.priorities.veryHigh;
  elements.priorityHigh.value = settings.theme.priorities.high;
  elements.priorityMedium.value = settings.theme.priorities.medium;
  elements.priorityLow.value = settings.theme.priorities.low;
  elements.priorityOptional.value = settings.theme.priorities.optional;
  updateFontScaleControl(settings.display?.fontScale || 1);
};

const updatePreview = () => {
  applyProfilePreview(draft);
  applyThemePreview(draft);
  applyFontScalePreview(draft);
};

const handleInputChange = (event) => {
  const { id, value } = event.target;
  switch (id) {
    case "profileName":
      draft = {
        ...draft,
        profile: { ...draft.profile, name: value },
      };
      break;
    case "profilePhoto":
      draft = {
        ...draft,
        profile: { ...draft.profile, photo: value },
      };
      break;
    case "accentColor":
      draft = {
        ...draft,
        theme: {
          ...draft.theme,
          accent: value,
          priorities: { ...draft.theme.priorities },
        },
      };
      break;
    case "priorityCritical":
    case "priorityVeryHigh":
    case "priorityHigh":
    case "priorityMedium":
    case "priorityLow":
    case "priorityOptional":
      draft = {
        ...draft,
        theme: {
          ...draft.theme,
          priorities: {
            ...draft.theme.priorities,
            [id.replace("priority", "").charAt(0).toLowerCase() + id.replace("priority", "").slice(1)]: value,
          },
        },
      };
      break;
    case "fontScale": {
      const percent = Number(value) || 100;
      const scale = clampFontScale(percent / 100);
      draft = {
        ...draft,
        display: { ...draft.display, fontScale: scale },
      };
      break;
    }
    default:
      break;
  }
  updatePreview();
};

const showStatus = (message, tone = "success") => {
  if (!elements.saveStatus) return;
  elements.saveStatus.textContent = message;
  elements.saveStatus.className = tone === "success" ? "text-sm text-emerald-600" : "text-sm text-rose-600";
  window.setTimeout(() => {
    if (elements.saveStatus.textContent === message) {
      elements.saveStatus.textContent = "";
    }
  }, 3500);
};

elements.form.addEventListener("submit", async (event) => {
  event.preventDefault();
  draft = normaliseSettings(draft);
  saveLocalSettings(draft);
  await saveSettingsRemote(draft);
  updatePreview();
  showStatus("Settings saved. Refresh your workspace to apply.", "success");
});

elements.form.addEventListener("input", handleInputChange);

elements.resetButton.addEventListener("click", async () => {
  draft = normaliseSettings(defaultSettings());
  populateForm(draft);
  updatePreview();
  saveLocalSettings(draft);
  await saveSettingsRemote(draft);
  showStatus("Settings reset to defaults.", "success");
});

elements.exportWorkspace?.addEventListener("click", handleExportWorkspaceClick);

populateForm(draft);
updatePreview();

fetchRemoteSettings().then((remote) => {
  if (remote) {
    draft = remote;
    populateForm(draft);
    updatePreview();
    saveLocalSettings(draft);
  }
});

const loadLocalArray = (key, fallback = []) => {
  try {
    const raw = window.localStorage.getItem(key);
    if (!raw) return fallback;
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : fallback;
  } catch {
    return fallback;
  }
};

const saveLocalArray = (key, value) => {
  try {
    window.localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // ignore localStorage issues
  }
};

const syncTeamLocal = () => {
  saveLocalArray(STORAGE_KEYS.members, teamMembers);
  saveLocalArray(STORAGE_KEYS.departments, teamDepartments);
};

const ensureDefaultDepartmentPresent = () => {
  if (!teamDepartments.some((department) => department.id === DEFAULT_DEPARTMENT.id)) {
    teamDepartments = [{ ...DEFAULT_DEPARTMENT }, ...teamDepartments];
  }
};

const getDepartmentName = (departmentId) =>
  teamDepartments.find((department) => department.id === departmentId)?.name || "Unassigned";

const populateMemberDepartmentOptions = (selectedId = "") => {
  if (!elements.memberDepartment) return;
  const fragment = document.createDocumentFragment();
  const defaultOption = document.createElement("option");
  defaultOption.value = "";
  defaultOption.textContent = "No department";
  fragment.append(defaultOption);
  teamDepartments.forEach((department) => {
    const option = document.createElement("option");
    option.value = department.id;
    option.textContent = department.name;
    option.disabled = Boolean(department.isDefault && teamDepartments.length === 1);
    fragment.append(option);
  });
  elements.memberDepartment.replaceChildren(fragment);
  elements.memberDepartment.value = selectedId || "";
};

const createDepartmentSelect = (selectedId = "", dataset = {}) => {
  const select = document.createElement("select");
  select.className = "field-input inline-select";
  Object.entries(dataset).forEach(([key, value]) => {
    if (value === undefined) return;
    select.dataset[key] = value;
  });

  const defaultOption = document.createElement("option");
  defaultOption.value = "";
  defaultOption.textContent = "No department";
  select.append(defaultOption);

  teamDepartments.forEach((department) => {
    const option = document.createElement("option");
    option.value = department.id;
    option.textContent = department.name;
    select.append(option);
  });

  select.value = selectedId || "";
  return select;
};

const renderDepartmentList = () => {
  if (!elements.departmentList) return;
  if (!teamDepartments.length) {
    const empty = document.createElement("li");
    empty.textContent = "No departments yet.";
    elements.departmentList.replaceChildren(empty);
    return;
  }

  const fragment = document.createDocumentFragment();
  teamDepartments.forEach((department) => {
    const item = document.createElement("li");
    item.className = "settings-inline-item";

    const fields = document.createElement("div");
    fields.className = "settings-inline-fields";

    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.className = "field-input inline-input";
    nameInput.value = department.name;
    nameInput.dataset.departmentName = department.id;
    fields.append(nameInput);

    item.append(fields);

    const actions = document.createElement("div");
    actions.className = "settings-inline-actions";

    const saveButton = document.createElement("button");
    saveButton.type = "button";
    saveButton.className = "btn-secondary btn-compact";
    saveButton.dataset.action = "update-department";
    saveButton.dataset.departmentId = department.id;
    saveButton.textContent = "Save";
    actions.append(saveButton);

    if (!department.isDefault) {
      const removeButton = document.createElement("button");
      removeButton.type = "button";
      removeButton.className = "ghost-button small";
      removeButton.dataset.action = "remove-department";
      removeButton.dataset.departmentId = department.id;
      removeButton.textContent = "Remove";
      actions.append(removeButton);
    } else {
      const badge = document.createElement("span");
      badge.className = "default-badge";
      badge.textContent = "Default";
      actions.append(badge);
    }

    item.append(actions);
    fragment.append(item);
  });

  elements.departmentList.replaceChildren(fragment);
};

const renderMemberList = () => {
  if (!elements.memberList) return;
  if (!teamMembers.length) {
    const empty = document.createElement("li");
    empty.textContent = "No members yet.";
    elements.memberList.replaceChildren(empty);
    return;
  }

  const fragment = document.createDocumentFragment();
  teamMembers.forEach((member) => {
    const item = document.createElement("li");
    item.className = "settings-inline-item";

    const fields = document.createElement("div");
    fields.className = "settings-inline-fields";

    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.className = "field-input inline-input";
    nameInput.value = member.name;
    nameInput.dataset.memberName = member.id;
    fields.append(nameInput);

    const departmentSelect = createDepartmentSelect(member.departmentId, {
      memberDepartment: member.id,
    });
    fields.append(departmentSelect);

    item.append(fields);

    const actions = document.createElement("div");
    actions.className = "settings-inline-actions";

    const saveButton = document.createElement("button");
    saveButton.type = "button";
    saveButton.className = "btn-secondary btn-compact";
    saveButton.dataset.action = "update-member";
    saveButton.dataset.memberId = member.id;
    saveButton.textContent = "Save";
    actions.append(saveButton);

    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className = "ghost-button small";
    removeButton.dataset.action = "remove-member";
    removeButton.dataset.memberId = member.id;
    removeButton.textContent = "Remove";
    actions.append(removeButton);

    item.append(actions);
    fragment.append(item);
  });

  elements.memberList.replaceChildren(fragment);
};

const syncTeamRemote = async () => {
  try {
    await setDoc(
      workspaceRef,
      {
        members: teamMembers,
        departments: teamDepartments,
        updatedAt: new Date().toISOString(),
      },
      { merge: true },
    );
  } catch (error) {
    console.error('Failed to sync team data to Firestore', error);
  }
};

const addDepartment = (name) => {
  const trimmed = name.trim();
  if (!trimmed) return { error: "Department name is required." };
  const exists = teamDepartments.some((department) => department.name.toLowerCase() === trimmed.toLowerCase());
  if (exists) return { error: "That department already exists." };
  const department = {
    id: generateId("department"),
    name: trimmed,
    isDefault: false,
    createdAt: new Date().toISOString(),
  };
  teamDepartments.push(department);
  ensureDefaultDepartmentPresent();
  syncTeamLocal();
  populateMemberDepartmentOptions(elements.memberDepartment?.value || "");
  renderDepartmentList();
  renderMemberList();
  syncTeamRemote();
  return { department };
};

const updateDepartment = (departmentId, nextName) => {
  const department = teamDepartments.find((entry) => entry.id === departmentId);
  if (!department) return { error: "Department not found." };
  const trimmed = nextName.trim();
  if (!trimmed) return { error: "Department name is required." };
  const exists = teamDepartments.some(
    (entry) => entry.id !== departmentId && entry.name.toLowerCase() === trimmed.toLowerCase(),
  );
  if (exists) return { error: "That department already exists." };

  department.name = trimmed;
  department.updatedAt = new Date().toISOString();
  syncTeamLocal();
  renderDepartmentList();
  populateMemberDepartmentOptions(elements.memberDepartment?.value || "");
  renderMemberList();
  syncTeamRemote();
  return { success: true };
};

const removeDepartment = (departmentId) => {
  const department = teamDepartments.find((entry) => entry.id === departmentId);
  if (!department) return { error: "Department not found." };
  if (department.isDefault) return { error: "The default department cannot be removed." };
  teamDepartments = teamDepartments.filter((entry) => entry.id !== departmentId);
  teamMembers = teamMembers.map((member) =>
    member.departmentId === departmentId ? { ...member, departmentId: "" } : member,
  );
  ensureDefaultDepartmentPresent();
  syncTeamLocal();
  populateMemberDepartmentOptions(elements.memberDepartment?.value || "");
  renderDepartmentList();
  renderMemberList();
  syncTeamRemote();
  return { success: true };
};

const addMember = (name, departmentId) => {
  const trimmed = name.trim();
  if (!trimmed) return { error: "Member name is required." };
  const member = {
    id: generateId("member"),
    name: trimmed,
    departmentId: departmentId || "",
    createdAt: new Date().toISOString(),
  };
  teamMembers.push(member);
  syncTeamLocal();
  renderMemberList();
  syncTeamRemote();
  return { member };
};

const updateMember = (memberId, updates) => {
  const memberIndex = teamMembers.findIndex((member) => member.id === memberId);
  if (memberIndex === -1) return { error: "Member not found." };
  const nextName = (updates?.name ?? teamMembers[memberIndex].name).trim();
  if (!nextName) return { error: "Member name is required." };
  const nextDepartment = updates?.departmentId ?? teamMembers[memberIndex].departmentId ?? "";

  teamMembers[memberIndex] = {
    ...teamMembers[memberIndex],
    name: nextName,
    departmentId: nextDepartment,
    updatedAt: new Date().toISOString(),
  };

  syncTeamLocal();
  renderMemberList();
  populateMemberDepartmentOptions(elements.memberDepartment?.value || "");
  syncTeamRemote();
  return { success: true };
};

const removeMember = (memberId) => {
  teamMembers = teamMembers.filter((member) => member.id !== memberId);
  syncTeamLocal();
  renderMemberList();
  syncTeamRemote();
};

const loadTeamData = async () => {
  teamMembers = loadLocalArray(STORAGE_KEYS.members, []);
  teamDepartments = loadLocalArray(STORAGE_KEYS.departments, []);
  ensureDefaultDepartmentPresent();
  try {
    const snapshot = await getDoc(workspaceRef);
    if (snapshot.exists()) {
      const data = snapshot.data();
      if (Array.isArray(data.members)) {
        teamMembers = data.members;
      }
      if (Array.isArray(data.departments)) {
        teamDepartments = data.departments;
      }
      ensureDefaultDepartmentPresent();
      syncTeamLocal();
    }
  } catch (error) {
    console.error('Failed to load team data from Firestore', error);
  }
  ensureDefaultDepartmentPresent();
  populateMemberDepartmentOptions();
  renderDepartmentList();
  renderMemberList();
};

const handleMemberFormSubmit = (event) => {
  event.preventDefault();
  if (!elements.memberForm) return;
  const name = elements.memberForm.memberName.value;
  const departmentId = elements.memberDepartment?.value || "";
  const { error } = addMember(name, departmentId) || {};
  if (error) {
    if (elements.memberError) elements.memberError.textContent = error;
    return;
  }
  if (elements.memberError) elements.memberError.textContent = "";
  elements.memberForm.reset();
  populateMemberDepartmentOptions();
};

const handleMemberListClick = (event) => {
  const button = event.target.closest('[data-action]');
  if (!button) return;
  const action = button.dataset.action;
  if (action === "remove-member") {
    removeMember(button.dataset.memberId);
    if (elements.memberError) elements.memberError.textContent = "";
    return;
  }
  if (action === "update-member") {
    const item = button.closest("li");
    if (!item) return;
    const memberId = button.dataset.memberId;
    const nameInput = item.querySelector('input[data-member-name]');
    const departmentSelect = item.querySelector('select[data-member-department]');
    const { error } =
      updateMember(memberId, {
        name: nameInput?.value ?? "",
        departmentId: departmentSelect?.value ?? "",
      }) || {};
    if (elements.memberError) {
      elements.memberError.textContent = error ? error : "";
    }
  }
};

const handleDepartmentFormSubmit = (event) => {
  event.preventDefault();
  if (!elements.departmentForm) return;
  const name = elements.departmentForm.departmentName.value;
  const { error } = addDepartment(name) || {};
  if (error) {
    if (elements.departmentError) elements.departmentError.textContent = error;
    return;
  }
  if (elements.departmentError) elements.departmentError.textContent = "";
  elements.departmentForm.reset();
};

const handleDepartmentListClick = (event) => {
  const button = event.target.closest('[data-action]');
  if (!button) return;
  const action = button.dataset.action;
  if (action === "remove-department") {
    const { error } = removeDepartment(button.dataset.departmentId) || {};
    if (elements.departmentError) elements.departmentError.textContent = error ? error : "";
    return;
  }
  if (action === "update-department") {
    const item = button.closest("li");
    if (!item) return;
    const input = item.querySelector('input[data-department-name]');
    const { error } = updateDepartment(button.dataset.departmentId, input?.value ?? "") || {};
    if (elements.departmentError) elements.departmentError.textContent = error ? error : "";
  }
};

const handleExportWorkspaceClick = async () => {
  if (!elements.exportWorkspace || elements.exportWorkspace.disabled) return;
  elements.exportWorkspace.disabled = true;
  elements.exportWorkspace.setAttribute("aria-busy", "true");

  let remoteData = null;
  try {
    const snapshot = await getDoc(workspaceRef);
    if (snapshot.exists()) {
      remoteData = snapshot.data();
    }
  } catch (error) {
    console.error("Failed to refresh workspace snapshot before export", error);
  }

  const pickArray = (value, fallback = []) => (Array.isArray(value) ? value : fallback);

  const payload = {
    tasks: pickArray(remoteData?.tasks, readLocalJSON(STORAGE_KEYS.tasks, [])),
    projects: pickArray(remoteData?.projects, readLocalJSON(STORAGE_KEYS.projects, [])),
    sections: pickArray(remoteData?.sections, readLocalJSON(STORAGE_KEYS.sections, [])),
    companies: pickArray(remoteData?.companies, readLocalJSON(STORAGE_KEYS.companies, [])),
    members: pickArray(remoteData?.members, teamMembers),
    departments: pickArray(remoteData?.departments, teamDepartments),
    userguide: pickArray(remoteData?.userguide, readLocalJSON(STORAGE_KEYS.userguide, [])),
    exportedAt: new Date().toISOString(),
  };

  try {
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = "synergy-tasks-backup.json";
    anchor.rel = "noopener";
    anchor.click();
    URL.revokeObjectURL(url);
    showStatus("Workspace export downloaded.", "success");
  } catch (error) {
    console.error("Failed to export workspace snapshot", error);
    showStatus("Unable to export workspace right now.", "error");
  } finally {
    elements.exportWorkspace.disabled = false;
    elements.exportWorkspace.removeAttribute("aria-busy");
  }
};

const initTeamManagement = async () => {
  await loadTeamData();
  elements.memberForm?.addEventListener("submit", handleMemberFormSubmit);
  elements.memberList?.addEventListener("click", handleMemberListClick);
  elements.departmentForm?.addEventListener("submit", handleDepartmentFormSubmit);
  elements.departmentList?.addEventListener("click", handleDepartmentListClick);
};

initTeamManagement();
