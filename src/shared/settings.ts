import { DEFAULT_SETTINGS, type ActionId, type EmailTemplate, type Settings } from './config';

const KEY = 'settings';

/**
 * Merges stored settings over defaults. Templates are merged per-action so a
 * newly added template appears for existing users instead of being dropped.
 */
/** Shape of settings written by earlier versions. */
type LegacySettings = Partial<Settings> & { bookingUrl?: string };

/** Ignores blank-line spacers so wording can be compared independently of spacing. */
function words(html: string | undefined): string {
  return (html ?? '')
    .replace(/<p>(?:&nbsp;|\s)*<\/p>/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Decides whether a stored template body may be replaced by the shipped default.
 *
 * Stored settings win over defaults, which is right for anything the user wrote —
 * but it also means an improvement to a shipped template never reaches someone who
 * saved it earlier. That is why the blank-line spacers were missing from real
 * emails while the Options preview looked correct.
 *
 * Two cases are safe to upgrade, and nothing else is:
 *   1. the wording is identical and only the spacing differs;
 *   2. the stored copy was still flagged as an unreviewed placeholder.
 */
function mayUpgrade(stored: EmailTemplate | undefined, shipped: EmailTemplate): boolean {
  if (!stored?.bodyHtml) return true;
  if (words(stored.bodyHtml) === words(shipped.bodyHtml)) return true;
  if (stored.isPlaceholder === true && shipped.isPlaceholder === false) return true;
  return false;
}

export function mergeStoredSettings(stored: LegacySettings | undefined): Settings {
  const merged = { ...DEFAULT_SETTINGS, ...stored };
  merged.templates = { ...DEFAULT_SETTINGS.templates };

  for (const [id, template] of Object.entries(stored?.templates ?? {})) {
    const key = id as ActionId;
    // Drops templates whose action no longer exists (the single 'shortlist'
    // action became two), and picks up newly added ones.
    const shipped = DEFAULT_SETTINGS.templates[key];
    if (!shipped) continue;

    if (mayUpgrade(template as EmailTemplate, shipped)) {
      // Take the shipped copy, but never discard settings the user actually chose.
      merged.templates[key] = {
        ...shipped,
        ...(template.bookingUrl?.trim() ? { bookingUrl: template.bookingUrl } : {}),
        ...(template.interviewerName?.trim() ? { interviewerName: template.interviewerName } : {}),
        ...(template.interviewerTitle?.trim() ? { interviewerTitle: template.interviewerTitle } : {}),
      };
    } else {
      merged.templates[key] = { ...shipped, ...template };
    }
  }

  // Migration: the Bookings URL used to be one global setting. Carry a stored
  // value onto any shortlist template that has none of its own.
  const legacyUrl = stored?.bookingUrl?.trim();
  if (legacyUrl) {
    for (const key of Object.keys(merged.templates) as ActionId[]) {
      const template = merged.templates[key];
      if (template.requiresBookingUrl && !template.bookingUrl?.trim()) {
        merged.templates[key] = { ...template, bookingUrl: legacyUrl };
      }
    }
  }
  delete (merged as LegacySettings).bookingUrl;

  return merged;
}

export async function loadSettings(): Promise<Settings> {
  try {
    const stored = await chrome.storage.local.get(KEY);
    return mergeStoredSettings(stored[KEY] as Partial<Settings> | undefined);
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

export async function saveSettings(patch: Partial<Settings>): Promise<Settings> {
  const current = await loadSettings();
  const next = { ...current, ...patch };
  await chrome.storage.local.set({ [KEY]: next });
  return next;
}

/** Wipes stored settings so the shipped defaults apply again. */
export async function resetSettings(): Promise<void> {
  await chrome.storage.local.remove(KEY);
}

