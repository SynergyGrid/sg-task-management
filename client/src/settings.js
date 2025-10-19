import { db } from './lib/firebaseClient';
import { doc, getDoc, setDoc } from 'firebase/firestore';

const STORAGE_KEYS = {
  settings: "synergygrid.todoist.settings.v1",
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
};

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
