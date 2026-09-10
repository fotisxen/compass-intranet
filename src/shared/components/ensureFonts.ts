const FONT_LINK_ID = 'compass-inter-font';

export function ensureInterFont(): void {
  if (document.getElementById(FONT_LINK_ID)) {
    return;
  }
  const link = document.createElement('link');
  link.id = FONT_LINK_ID;
  link.rel = 'stylesheet';
  link.href = 'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap';
  document.head.appendChild(link);
}
