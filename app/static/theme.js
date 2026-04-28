const THEME_KEY = "russell-toolkit-theme";
const THEME_DARK = "dark";
const THEME_LIGHT = "light";

function resolveInitialTheme() {
  try {
    const storedTheme = localStorage.getItem(THEME_KEY);
    if (storedTheme === THEME_DARK || storedTheme === THEME_LIGHT) {
      return storedTheme;
    }
  } catch (_error) {}

  const prefersLight =
    window.matchMedia && window.matchMedia("(prefers-color-scheme: light)").matches;
  return prefersLight ? THEME_LIGHT : THEME_DARK;
}

function applyTheme(theme) {
  document.documentElement.setAttribute("data-theme", theme);
}

function updateToggleCopy(button, theme) {
  button.textContent = theme === THEME_DARK ? "Light Mode" : "Dark Mode";
}

function initThemeToggle() {
  const toggle = document.getElementById("theme-toggle");
  if (!toggle) {
    return;
  }

  let currentTheme = resolveInitialTheme();
  applyTheme(currentTheme);
  updateToggleCopy(toggle, currentTheme);

  toggle.addEventListener("click", () => {
    currentTheme = currentTheme === THEME_DARK ? THEME_LIGHT : THEME_DARK;
    applyTheme(currentTheme);
    updateToggleCopy(toggle, currentTheme);
    try {
      localStorage.setItem(THEME_KEY, currentTheme);
    } catch (_error) {}
  });
}

document.addEventListener("DOMContentLoaded", initThemeToggle);
