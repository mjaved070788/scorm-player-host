/*
 * Player host configuration.
 *
 * ALLOWED_PARENT_ORIGINS: pages that may load SCORM packages into this player.
 * Each entry is an exact origin or a regular expression (string starting with ^).
 */
window.SCORM_PLAYER_CONFIG = {
  ALLOWED_PARENT_ORIGINS: [
    // TISC Digitalization SharePoint (SharePoint Online on the company domain)
    'https://groups.coca-cola.com',
    // Microsoft Teams (if the web part is added as a Teams tab)
    'https://teams.microsoft.com',
    // Local testing only - remove in production if you like
    '^https?://localhost(:\\d+)?$'
  ],
  // Packages cached by the service worker are removed after this many days.
  CACHE_MAX_AGE_DAYS: 14
};
