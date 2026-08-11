import { loadSettings, saveSettings, resetSettings } from '../shared/settings';
import {
  ACTION_ORDER,
  DEFAULT_SETTINGS,
  type ActionId,
  type EmailTemplate,
  type Settings,
} from '../shared/config';
import { renderTemplate, validateBookingUrl } from '../shared/template';

/** Fictional sample used for previews so no real candidate appears here. */
const PREVIEW_CANDIDATE = { email: 'candidate@example.com' };

const TEXT_FIELDS = [
  'jobTitle',
  'companyName',
  'senderName',
  'senderTitle',
  'greetingFallback',
] as const satisfies ReadonlyArray<keyof Settings>;

const TOGGLES = [
  'attemptSourceDeletion',
  'requireDeletionConfirmation',
  'debugLogging',
] as const satisfies ReadonlyArray<keyof Settings>;

let settings: Settings;

/* ------------------------------------------------------------------ */
/* helpers                                                            */
/* ------------------------------------------------------------------ */

function saved(): void {
  const node = document.getElementById('status');
  if (!node) return;
  node.textContent = 'Saved';
  node.classList.add('show');
  window.setTimeout(() => node.classList.remove('show'), 1600);
}

function debounce(fn: () => void, ms = 300): () => void {
  let timer: number | undefined;
  return () => {
    if (timer) window.clearTimeout(timer);
    timer = window.setTimeout(fn, ms);
  };
}

function stateLine(kind: 'ok' | 'warn' | 'err', text: string): HTMLElement {
  const el = document.createElement('p');
  el.className = `state ${kind}`;
  el.textContent = text;
  return el;
}

async function persist(patch: Partial<Settings>): Promise<void> {
  settings = await saveSettings(patch);
  saved();
  refreshAllPreviews();
}

async function saveTemplate(id: ActionId, patch: Partial<EmailTemplate>): Promise<void> {
  await persist({ templates: { ...settings.templates, [id]: { ...settings.templates[id], ...patch } } });
}

/* ------------------------------------------------------------------ */
/* preview                                                            */
/* ------------------------------------------------------------------ */

function refreshPreview(id: ActionId): void {
  const host = document.querySelector(`[data-preview="${id}"]`);
  if (!host) return;
  host.replaceChildren();

  const template = settings.templates[id];
  const result = renderTemplate(template, PREVIEW_CANDIDATE, settings);

  const frame = document.createElement('div');
  frame.className = 'preview';
  const subject = document.createElement('div');
  subject.className = 'preview-subject';
  subject.textContent = result.subject || '(no subject)';
  const body = document.createElement('div');
  body.className = 'preview-body';
  // The user's own template markup. Extension-page CSP blocks inline handlers
  // and remote loads, so this cannot execute anything.
  body.innerHTML = result.bodyHtml;
  frame.append(subject, body);
  host.append(frame);

  for (const error of result.errors) host.append(stateLine('err', `Blocked — ${error}`));
  for (const warning of result.warnings) host.append(stateLine('warn', warning));

  if (template.requiresBookingUrl) {
    const raw = template.bookingUrl ?? '';
    if (!raw.trim()) {
      host.append(stateLine('warn', 'No Bookings link set — this action cannot be used.'));
    } else {
      const { url, error, warning } = validateBookingUrl(raw);
      if (error) host.append(stateLine('err', error));
      else if (warning) host.append(stateLine('warn', warning));
      else if (url) host.append(stateLine('ok', `Bookings link valid (${new URL(url).host}).`));
    }
  }

  if (result.ok && result.errors.length === 0 && result.warnings.length === 0) {
    host.append(stateLine('ok', 'Ready to use.'));
  }
}

function refreshAllPreviews(): void {
  for (const id of ACTION_ORDER) refreshPreview(id);
}

/* ------------------------------------------------------------------ */
/* template editors                                                   */
/* ------------------------------------------------------------------ */

function labelledField(labelText: string, control: HTMLElement, help?: string): HTMLElement {
  const row = document.createElement('div');
  row.className = 'field';
  const label = document.createElement('label');
  label.textContent = labelText;
  const wrap = document.createElement('div');
  wrap.className = 'control';
  wrap.append(control);
  if (help) {
    const hint = document.createElement('p');
    hint.className = 'help';
    hint.textContent = help;
    wrap.append(hint);
  }
  const id = `f-${Math.abs(hashCode(labelText + labelText.length))}`;
  control.id = control.id || id;
  label.htmlFor = control.id;
  row.append(label, wrap);
  return row;
}

function hashCode(value: string): number {
  let hash = 0;
  for (let i = 0; i < value.length; i += 1) hash = (hash * 31 + value.charCodeAt(i)) | 0;
  return hash;
}

function buildTemplateEditors(): void {
  const list = document.getElementById('templates-list');
  if (!list) return;
  list.replaceChildren();

  for (const id of ACTION_ORDER) {
    const template = settings.templates[id];

    const details = document.createElement('details');
    details.className = `tpl ${template.tone === 'positive' ? 'pos' : 'neg'}`;
    details.open = template.isPlaceholder;

    const summary = document.createElement('summary');
    const chev = document.createElement('span');
    chev.className = 'chev';
    chev.textContent = '▶';
    const name = document.createElement('span');
    name.className = 'tname';
    name.textContent = template.label;
    const desc = document.createElement('span');
    desc.className = 'tdesc';
    desc.textContent = template.description;
    summary.append(chev, name, desc);
    if (template.isPlaceholder) {
      const badge = document.createElement('span');
      badge.className = 'badge';
      badge.textContent = 'needs review';
      summary.append(badge);
    }
    details.append(summary);

    const body = document.createElement('div');
    body.className = 'tpl-body';

    const subjectInput = document.createElement('input');
    subjectInput.type = 'text';
    subjectInput.id = `subject-${id}`;
    subjectInput.value = template.subject;
    body.append(labelledField('Subject', subjectInput));

    if (template.requiresBookingUrl) {
      const interviewer = document.createElement('input');
      interviewer.type = 'text';
      interviewer.id = `interviewer-${id}`;
      interviewer.value = template.interviewerName ?? '';
      body.append(labelledField('Interviewer', interviewer, 'Replaces {{INTERVIEWER}} in the body.'));

      const booking = document.createElement('input');
      booking.type = 'url';
      booking.id = `booking-${id}`;
      booking.value = template.bookingUrl ?? '';
      booking.placeholder = 'https://outlook.office.com/bookwithme/user/…';
      body.append(
        labelledField(
          'Bookings link',
          booking,
          'This action books this calendar. Must be https.',
        ),
      );

      const interviewerTitle = document.createElement('input');
      interviewerTitle.type = 'text';
      interviewerTitle.id = `interviewer-title-${id}`;
      interviewerTitle.value = template.interviewerTitle ?? '';
      body.append(
        labelledField('Interviewer role', interviewerTitle, 'Replaces {{INTERVIEWER_TITLE}}.'),
      );
      interviewerTitle.addEventListener(
        'input',
        debounce(() => void saveTemplate(id, { interviewerTitle: interviewerTitle.value })),
      );

      interviewer.addEventListener(
        'input',
        debounce(() => void saveTemplate(id, { interviewerName: interviewer.value })),
      );
      booking.addEventListener(
        'input',
        debounce(() => void saveTemplate(id, { bookingUrl: booking.value })),
      );
    }

    const bodyArea = document.createElement('textarea');
    bodyArea.id = `body-${id}`;
    bodyArea.rows = 14;
    bodyArea.value = template.bodyHtml.replace(/\n/g, '\n');
    body.append(labelledField('Body (HTML)', bodyArea));

    const actions = document.createElement('div');
    actions.className = 'tpl-actions';

    if (template.isPlaceholder) {
      const reviewed = document.createElement('button');
      reviewed.type = 'button';
      reviewed.className = 'btn primary';
      reviewed.textContent = 'Mark as reviewed';
      reviewed.addEventListener('click', () => {
        void saveTemplate(id, { isPlaceholder: false }).then(buildTemplateEditors);
      });
      actions.append(reviewed);
    }

    const spacer = document.createElement('button');
    spacer.type = 'button';
    spacer.className = 'btn';
    spacer.textContent = 'Insert blank line';
    spacer.title = 'Inserts <p>&nbsp;</p> at the cursor — a blank line between paragraphs';
    spacer.addEventListener('click', () => {
      const at = bodyArea.selectionStart ?? bodyArea.value.length;
      const insert = `${at > 0 && bodyArea.value[at - 1] !== '\n' ? '\n' : ''}<p>&nbsp;</p>\n`;
      bodyArea.value = bodyArea.value.slice(0, at) + insert + bodyArea.value.slice(at);
      bodyArea.focus();
      bodyArea.selectionStart = bodyArea.selectionEnd = at + insert.length;
      void saveTemplate(id, { bodyHtml: bodyArea.value });
    });
    actions.append(spacer);

    const reset = document.createElement('button');
    reset.type = 'button';
    reset.className = 'btn';
    reset.textContent = 'Reset to default';
    reset.addEventListener('click', () => {
      const original = DEFAULT_SETTINGS.templates[id];
      void saveTemplate(id, {
        subject: original.subject,
        bodyHtml: original.bodyHtml,
        isPlaceholder: original.isPlaceholder,
        bookingUrl: original.bookingUrl,
        interviewerName: original.interviewerName,
        interviewerTitle: original.interviewerTitle,
      }).then(buildTemplateEditors);
    });
    actions.append(reset);
    body.append(actions);

    const previewLabel = document.createElement('p');
    previewLabel.className = 'preview-label';
    previewLabel.textContent = 'Preview';
    const previewHost = document.createElement('div');
    previewHost.dataset.preview = id;
    body.append(previewLabel, previewHost);

    details.append(body);
    list.append(details);

    subjectInput.addEventListener(
      'input',
      debounce(() => void saveTemplate(id, { subject: subjectInput.value })),
    );
    bodyArea.addEventListener(
      'input',
      debounce(() => void saveTemplate(id, { bodyHtml: bodyArea.value })),
    );
  }

  refreshAllPreviews();
}

/* ------------------------------------------------------------------ */
/* init                                                              */
/* ------------------------------------------------------------------ */

async function init(): Promise<void> {
  settings = await loadSettings();

  const versionEl = document.getElementById('version');
  if (versionEl) {
    try {
      versionEl.textContent = chrome.runtime.getManifest().version;
    } catch {
      /* running outside the extension, e.g. the harness */
    }
  }

  for (const key of TEXT_FIELDS) {
    const input = document.getElementById(key) as HTMLInputElement | null;
    if (!input) continue;
    input.value = String(settings[key] ?? '');
    input.addEventListener(
      'input',
      debounce(() => void persist({ [key]: input.value } as Partial<Settings>)),
    );
  }

  for (const key of TOGGLES) {
    const input = document.getElementById(key) as HTMLInputElement | null;
    if (!input) continue;
    input.checked = Boolean(settings[key]);
    input.addEventListener('change', () => {
      void persist({ [key]: input.checked } as Partial<Settings>);
    });
  }

  const greeting = document.getElementById('greetingMode') as HTMLSelectElement | null;
  if (greeting) {
    greeting.value = settings.greetingMode;
    greeting.addEventListener('change', () => {
      void persist({ greetingMode: greeting.value as Settings['greetingMode'] });
    });
  }

  const resetAll = document.getElementById('resetAll');
  resetAll?.addEventListener('click', () => {
    if (!window.confirm('Reset every setting and template to its shipped default?')) return;
    void resetSettings().then(async () => {
      settings = await loadSettings();
      buildTemplateEditors();
      void init();
      saved();
    });
  });

  buildTemplateEditors();
}

void init();
