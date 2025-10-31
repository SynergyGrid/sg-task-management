import { db } from './lib/firebaseClient';
import { doc, getDoc, setDoc } from 'firebase/firestore';

const STORAGE_KEYS = {
  settings: "synergygrid.todoist.settings.v1",
  members: "synergygrid.todoist.members.v1",
  departments: "synergygrid.todoist.departments.v1",
};

const WORKSPACE_ID = import.meta.env.VITE_FIREBASE_WORKSPACE_ID ?? 'default';
const workspaceRef = doc(db, 'workspaces', WORKSPACE_ID);

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
};

const updatePreview = () => {
  applyProfilePreview(draft);
  applyThemePreview(draft);
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

const renderDepartmentList = () => {
  if (!elements.departmentList) return;
  const fragment = document.createDocumentFragment();
  teamDepartments.forEach((department) => {
    const item = document.createElement("li");
    const label = document.createElement("span");
    label.textContent = department.name;
    item.append(label);
    if (!department.isDefault) {
      const removeButton = document.createElement("button");
      removeButton.type = "button";
      removeButton.className = "ghost-button small";
      removeButton.dataset.action = "remove-department";
      removeButton.dataset.departmentId = department.id;
      removeButton.textContent = "Remove";
      item.append(removeButton);
    }
    fragment.append(item);
  });
  if (!fragment.childElementCount) {
    const empty = document.createElement("li");
    empty.textContent = "No departments yet.";
    fragment.append(empty);
  }
  elements.departmentList.replaceChildren(fragment);
};

const renderMemberList = () => {
  if (!elements.memberList) return;
  const fragment = document.createDocumentFragment();
  teamMembers.forEach((member) => {
    const item = document.createElement("li");
    const info = document.createElement("div");
    const name = document.createElement("span");
    name.className = "font-medium";
    name.textContent = member.name;
    const meta = document.createElement("small");
    meta.className = "text-slate-500";
    meta.textContent = getDepartmentName(member.departmentId);
    info.append(name, meta);
    item.append(info);
    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className = "ghost-button small";
    removeButton.dataset.action = "remove-member";
    removeButton.dataset.memberId = member.id;
    removeButton.textContent = "Remove";
    item.append(removeButton);
    fragment.append(item);
  });
  if (!fragment.childElementCount) {
    const empty = document.createElement("li");
    empty.textContent = "No members yet.";
    fragment.append(empty);
  }
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
  if (button.dataset.action === "remove-member") {
    removeMember(button.dataset.memberId);
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
  if (button.dataset.action === "remove-department") {
    const { error } = removeDepartment(button.dataset.departmentId) || {};
    if (error) {
      if (elements.departmentError) elements.departmentError.textContent = error;
    } else if (elements.departmentError) {
      elements.departmentError.textContent = "";
    }
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
