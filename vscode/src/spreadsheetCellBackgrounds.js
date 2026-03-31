"use strict";

const BACKGROUND_COLOR_OPTIONS = [
  { token: "", label: "No color", css: "", argb: "" },
  { token: "red", label: "Red", css: "#f4c7c3", argb: "FFF4C7C3" },
  { token: "green", label: "Green", css: "#cfe8cc", argb: "FFCFE8CC" },
  { token: "blue", label: "Blue", css: "#c9daf8", argb: "FFC9DAF8" },
  { token: "yellow", label: "Yellow", css: "#fff2cc", argb: "FFFFF2CC" },
  { token: "magenta", label: "Magenta", css: "#ead1dc", argb: "FFEAD1DC" },
  { token: "cyan", label: "Cyan", css: "#d0e0e3", argb: "FFD0E0E3" },
  { token: "pink", label: "Pink", css: "#f4cccc", argb: "FFF4CCCC" },
  { token: "orange", label: "Orange", css: "#f9d5b4", argb: "FFF9D5B4" },
  { token: "brown", label: "Brown", css: "#d9c2a7", argb: "FFD9C2A7" },
  { token: "grey", label: "Grey", css: "#d9d9d9", argb: "FFD9D9D9" },
];

const BACKGROUND_OPTION_BY_TOKEN = new Map(
  BACKGROUND_COLOR_OPTIONS.map((option) => [option.token, option]),
);

const BACKGROUND_TOKEN_BY_ARGB = new Map(
  BACKGROUND_COLOR_OPTIONS
    .filter((option) => option.argb)
    .map((option) => [option.argb, option.token]),
);

function normalizeBackgroundToken(token) {
  const normalizedToken = String(token || "").trim().toLowerCase();
  return BACKGROUND_OPTION_BY_TOKEN.has(normalizedToken)
    ? normalizedToken
    : "";
}

function normalizeArgb(value) {
  const normalized = String(value || "").trim().toUpperCase();
  if (!normalized) {
    return "";
  }
  if (/^[0-9A-F]{6}$/.test(normalized)) {
    return `FF${normalized}`;
  }
  return /^[0-9A-F]{8}$/.test(normalized)
    ? normalized
    : "";
}

function getBackgroundCss(token) {
  return BACKGROUND_OPTION_BY_TOKEN.get(normalizeBackgroundToken(token))?.css || "";
}

function getBackgroundArgb(token) {
  return BACKGROUND_OPTION_BY_TOKEN.get(normalizeBackgroundToken(token))?.argb || "";
}

function getBackgroundTokenByArgb(argb) {
  return BACKGROUND_TOKEN_BY_ARGB.get(normalizeArgb(argb)) || "";
}

module.exports = {
  BACKGROUND_COLOR_OPTIONS,
  getBackgroundArgb,
  getBackgroundCss,
  getBackgroundTokenByArgb,
  normalizeArgb,
  normalizeBackgroundToken,
};
