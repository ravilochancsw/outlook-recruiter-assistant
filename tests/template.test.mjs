import test from 'node:test';
import assert from 'node:assert/strict';
import { lib } from './dom-setup.mjs';

const { renderTemplate, validateBookingUrl, escapeHtml, htmlToText, DEFAULT_SETTINGS, ACTION_ORDER } =
  await lib();

const CANDIDATE = { email: 'chitranshi@example.com' };

function settings(patch = {}) {
  return { ...DEFAULT_SETTINGS, ...patch };
}

function render(id, patch = {}, candidate = CANDIDATE) {
  const s = settings(patch);
  return renderTemplate(s.templates[id], candidate, s);
}

/** Overrides a single template field, e.g. its Bookings URL. */
function renderWith(id, templatePatch, patch = {}, candidate = CANDIDATE) {
  const s = settings(patch);
  return renderTemplate({ ...s.templates[id], ...templatePatch }, candidate, s);
}

const SHORTLIST_IDS = ['shortlist-ravilochan', 'shortlist-abhi'];
/** Every action that books a calendar, screenings plus the round-two interview. */
const BOOKING_IDS = [...SHORTLIST_IDS, 'interview-ravilochan'];

/* ------------------------------------------------------------------ */
/* the user's confirmed copy                                           */
/* ------------------------------------------------------------------ */

test('direct rejection renders the user’s confirmed copy', () => {
  const out = render('reject-direct');
  assert.equal(out.ok, true, out.errors.join('; '));
  assert.equal(out.subject, 'Update on Your Application');
  assert.match(out.bodyText, /^Thank you for your interest in AI Full Stack Software Engineer position at Cloud Security Web and for taking the time to apply\./);
  assert.match(out.bodyText, /we wish you all the best in your job search\./);
  assert.match(out.bodyText, /Best regards,\nRavilochan\nHead of Engineering/);
});

test('direct rejection keeps the bold on job title and company', () => {
  const out = render('reject-direct');
  assert.match(out.bodyHtml, /<strong>AI Full Stack Software Engineer<\/strong>/);
  assert.match(out.bodyHtml, /<strong>Cloud Security Web<\/strong>/);
});

test('in-process rejection renders the user’s confirmed copy, unbolded', () => {
  const out = render('reject-in-process');
  assert.equal(out.ok, true, out.errors.join('; '));
  assert.equal(out.subject, 'Update in Interview Process');
  assert.match(out.bodyText, /^Thank you for taking the time to apply for the AI Full Stack Software Engineer position at Cloud Security Web\./);
  assert.match(out.bodyText, /will keep your application on file/);
  assert.match(out.bodyText, /We wish you continued success in your job search\./);
  assert.ok(!out.bodyHtml.includes('<strong>'), 'this template has no bold in the original');
});

test('neither rejection template contains a greeting line', () => {
  for (const id of ['reject-direct', 'reject-in-process']) {
    const out = render(id);
    assert.ok(!/^\s*(hi|hello|dear)\b/i.test(out.bodyText), `${id} unexpectedly opens with a greeting`);
  }
});

test('rejection templates never need a Bookings URL', () => {
  for (const id of ['reject-direct', 'reject-in-process']) {
    const out = render(id);
    assert.equal(out.ok, true, `${id}: ${out.errors.join('; ')}`);
    assert.ok(!out.bodyHtml.includes('href'));
  }
});

/* ------------------------------------------------------------------ */
/* shortlist + Bookings URL                                            */
/* ------------------------------------------------------------------ */

test('every booking action ships with no configured Bookings link, and fails closed', () => {
  // Real links are per-installation (set in Options → Templates), never shipped —
  // shipping one would land a real mailbox and calendar in version control.
  for (const id of BOOKING_IDS) {
    assert.equal(DEFAULT_SETTINGS.templates[id].bookingUrl, undefined, `${id} ships a default link`);
    const out = render(id);
    assert.equal(out.ok, false);
    assert.match(out.errors.join(' '), /No Microsoft Bookings URL is configured/);
    // The message names the action, so the panel can say which button failed.
    assert.match(out.errors.join(' '), new RegExp(DEFAULT_SETTINGS.templates[id].label.replace(/[().]/g, '\\$&')));
  }
});

test('once configured, the two shortlist actions route to different calendars', () => {
  const raviUrl = 'https://outlook.office.com/bookwithme/user/ravi-fixture/meetingtype/one';
  const abhiUrl = 'https://outlook.office.com/bookwithme/user/abhi-fixture/meetingtype/two';
  const ravi = renderWith('shortlist-ravilochan', { bookingUrl: raviUrl });
  const abhi = renderWith('shortlist-abhi', { bookingUrl: abhiUrl });
  assert.equal(ravi.ok, true, ravi.errors.join('; '));
  assert.equal(abhi.ok, true, abhi.errors.join('; '));
  const link = (out) => out.bodyHtml.match(/href="([^"]+)"/)[1];
  assert.equal(link(ravi), raviUrl);
  assert.equal(link(abhi), abhiUrl);
  assert.notEqual(link(ravi), link(abhi));
});

test('a shortlist action is blocked when its configured Bookings URL is cleared', () => {
  for (const id of SHORTLIST_IDS) {
    const configured = renderWith(id, { bookingUrl: 'https://outlook.office.com/bookwithme/user/fixture' });
    assert.equal(configured.ok, true, configured.errors.join('; '));
    const cleared = renderWith(id, { bookingUrl: '' });
    assert.equal(cleared.ok, false);
    assert.match(cleared.errors.join(' '), /No Microsoft Bookings URL is configured/);
  }
});

test('clearing one shortlist link does not block the other', () => {
  const s = settings();
  const broken = renderTemplate({ ...s.templates['shortlist-ravilochan'], bookingUrl: '' }, CANDIDATE, s);
  const intact = renderTemplate(
    { ...s.templates['shortlist-abhi'], bookingUrl: 'https://outlook.office.com/bookwithme/user/fixture' },
    CANDIDATE,
    s,
  );
  assert.equal(broken.ok, false);
  assert.equal(intact.ok, true, intact.errors.join('; '));
});

test('a configured Bookings URL survives validation byte-for-byte', () => {
  // Real links carry valueless query params (e.g. &anonymous&ismsaljsauthenabled).
  // If URL normalisation dropped or reordered them, booking would break.
  const withTrickyParams =
    'https://outlook.office.com/bookwithme/user/fixture@example.com/meetingtype/abc?bookingcode=fixture-code&anonymous&ismsaljsauthenabled&ep=mlink';
  const { url, error } = validateBookingUrl(withTrickyParams);
  assert.equal(error, undefined, error);
  assert.equal(url, withTrickyParams, 'URL was rewritten');
});

test('distinct configured links for each booking action stay distinct after rendering', () => {
  const urls = {
    'shortlist-ravilochan': 'https://outlook.office.com/bookwithme/user/fixture/meetingtype/one',
    'shortlist-abhi': 'https://outlook.office.com/bookwithme/user/fixture/meetingtype/two',
    'interview-ravilochan': 'https://outlook.office.com/bookwithme/user/fixture/meetingtype/three',
  };
  const link = (id) => renderWith(id, { bookingUrl: urls[id] }).bodyHtml.match(/href="([^"]+)"/)[1];
  const all = BOOKING_IDS.map(link);
  assert.equal(new Set(all).size, all.length, 'two actions produced the same link from different config');
});

test('the round-two interview does not promise there is no coding exercise', () => {
  // It asks real technical questions. Promising otherwise would mislead the
  // candidate and skew the conversation.
  const out = render('interview-ravilochan');
  assert.ok(!/no coding|no technical (test|assessment)|nothing to prepare/i.test(out.bodyText));
});

test('the round-two interview sets up depth, design and a walkthrough', () => {
  const out = renderWith('interview-ravilochan', {
    bookingUrl: 'https://outlook.office.com/bookwithme/user/fixture',
  });
  assert.equal(out.ok, true, out.errors.join('; '));
  assert.equal(out.subject, 'Technical Interview – AI Full Stack Software Engineer');
  assert.match(out.bodyText, /one-hour technical discussion over Microsoft Teams/);
  assert.match(out.bodyText, /a walkthrough of something you’ve built/);
  assert.match(out.bodyText, /how you would design and structure a system/);
  assert.match(out.bodyText, /talk through end to end, including the parts that gave you trouble/);
  assert.equal((out.bodyHtml.match(/<li>/g) ?? []).length, 3);
});

test('the round-two interview reveals nothing about what follows it', () => {
  const out = render('interview-ravilochan');
  assert.ok(
    !/founder|final round|next step|next round|last stage/i.test(out.bodyText),
    'the interview email leaks the rest of the process',
  );
});

test('the interview and the screening are clearly different emails', () => {
  const screening = render('shortlist-ravilochan');
  const interview = render('interview-ravilochan');
  assert.notEqual(screening.subject, interview.subject);
  assert.match(screening.bodyText, /20-minute/);
  assert.match(interview.bodyText, /one-hour/);
});

test('ampersands are entity-escaped in href but raw in the text rendering', () => {
  // Real Bookings links carry valueless query params chained with bare &, e.g.
  // ?anonymous&ismsaljsauthenabled&ep=mcard — fixture reproduces that shape.
  const out = renderWith('shortlist-abhi', {
    bookingUrl: 'https://outlook.office.com/bookwithme/user/fixture?anonymous&ismsaljsauthenabled&ep=mcard',
  });
  assert.equal(out.ok, true, out.errors.join('; '));
  // ?anonymous is the first param; the separators before the rest must be escaped.
  assert.ok(out.bodyHtml.includes('&amp;ismsaljsauthenabled'), 'href must escape & to be valid HTML');
  assert.ok(out.bodyHtml.includes('&amp;ep=mcard'));
  assert.ok(!/&(?!amp;|lt;|gt;|quot;|#39;|nbsp;)/.test(out.bodyHtml), 'unescaped bare & in the body');
  // htmlToText must decode them again so the plain-text URL is clickable.
  assert.ok(out.bodyText.includes('?anonymous&ismsaljsauthenabled&ep=mcard'));
});

test('every shipped template is confirmed copy, not a placeholder', () => {
  for (const id of ACTION_ORDER) {
    // isPlaceholder is about the wording, independent of whether a Bookings
    // link happens to be configured yet — so it is checked directly rather
    // than through render(), which fails closed on the booking actions until
    // a link is set.
    assert.equal(DEFAULT_SETTINGS.templates[id].isPlaceholder, false, `${id} is still a placeholder`);

    const out = DEFAULT_SETTINGS.templates[id].requiresBookingUrl
      ? renderWith(id, { bookingUrl: 'https://outlook.office.com/bookwithme/user/fixture' })
      : render(id);
    assert.equal(out.ok, true, `${id}: ${out.errors.join('; ')}`);
    assert.deepEqual(out.warnings, [], `${id} still warns: ${out.warnings.join('; ')}`);
  }
});

test('a template marked as a placeholder still warns', () => {
  const out = renderWith('shortlist-abhi', { isPlaceholder: true });
  assert.match(out.warnings.join(' '), /placeholder/i);
});

test('the self-run screening is written in the first person', () => {
  const out = render('shortlist-ravilochan');
  assert.match(out.bodyText, /I’ve reviewed your application/);
  assert.match(out.bodyText, /I’d like to spend the time on what you’ve actually built/);
  assert.ok(!/with Abhi/.test(out.bodyText));
});

test('both screening templates set the expectation of technical depth', () => {
  for (const id of SHORTLIST_IDS) {
    const out = render(id);
    assert.match(out.bodyText, /Rather than going through your resume line by line/);
    assert.match(out.bodyText, /20-minute technical screening over Microsoft Teams/);
    assert.match(out.bodyText, /no coding exercise/);
    assert.match(out.bodyText, /Pick one or two projects you know well/);
    assert.match(out.bodyText, /one hour before they start/);
  }
});

test('the screening asks the three things a candidate cannot rehearse', () => {
  for (const id of SHORTLIST_IDS) {
    const out = render(id);
    assert.match(out.bodyText, /the part you worked on yourself/);
    assert.match(out.bodyText, /why you built it that way, and what you considered instead/);
    assert.match(out.bodyText, /what you would do differently now/);
  }
});

test('the screening reveals nothing about later rounds', () => {
  // Decided deliberately: what comes next is disclosed only if the candidate
  // progresses, or answered live if they ask on the call.
  for (const id of SHORTLIST_IDS) {
    const out = render(id);
    assert.ok(
      !/next step|next round|further round|longer (session|conversation|technical)|founder|final round|second (round|conversation)/i.test(
        out.bodyText,
      ),
      `${id} leaks the rest of the process`,
    );
  }
});

test('the screening carries no HR filler and no seniority signalling', () => {
  for (const id of SHORTLIST_IDS) {
    const out = render(id);
    assert.ok(!/drew you to|why you.{0,15}interested|attracted you/i.test(out.bodyText), 'HR filler');
    assert.ok(
      !/busy|limited time|my time|briefly as possible/i.test(out.bodyText),
      'seniority signalling makes candidates defensive and costs signal',
    );
  }
});

test('the subject stays the conventional screening wording', () => {
  for (const id of SHORTLIST_IDS) {
    assert.equal(render(id).subject, 'Technical Screening – AI Full Stack Software Engineer');
  }
});

test('the delegated screening names the interviewer, their role, and the contact', () => {
  const out = render('shortlist-abhi');
  assert.match(out.bodyText, /with Abhi, Software Engineer on our engineering team/);
  assert.match(out.bodyText, /Abhi will be your point of contact for this conversation/);
  // Still sent from the account holder.
  assert.match(out.bodyText, /Best regards,\nRavilochan\nHead of Engineering/);
});

test('both screening versions run the same agenda', () => {
  const [self, delegated] = SHORTLIST_IDS.map((id) => render(id).bodyText);
  for (const shared of [
    'Rather than going through your resume line by line',
    'the part you worked on yourself',
    'no coding exercise',
    'one hour before they start',
  ]) {
    assert.ok(self.includes(shared) && delegated.includes(shared), `diverged on: ${shared}`);
  }
});

test('{{INTERVIEWER_TITLE}} blocks rendering when unset on a template that uses it', () => {
  const out = renderWith('shortlist-abhi', { interviewerTitle: '' });
  assert.equal(out.ok, false);
  assert.match(out.errors.join(' '), /\{\{INTERVIEWER_TITLE\}\} has no value configured/);
});

test('the screening topics render as a real list of three', () => {
  for (const id of SHORTLIST_IDS) {
    const out = render(id);
    assert.match(out.bodyHtml, /<ul>/);
    assert.equal((out.bodyHtml.match(/<li>/g) ?? []).length, 3);
  }
});

test('{{INTERVIEWER}} fails rendering when no interviewer is set', () => {
  const out = renderWith('shortlist-abhi', { interviewerName: '' });
  assert.equal(out.ok, false);
  assert.match(out.errors.join(' '), /\{\{INTERVIEWER\}\} has no value configured/);
});

test('Bookings URL validation rejects non-https and malformed input', () => {
  assert.match(validateBookingUrl('http://example.com/x').error, /must use https/);
  assert.match(validateBookingUrl('not a url').error, /not a valid URL/);
  assert.match(validateBookingUrl('   ').error, /No Microsoft Bookings URL/);
  assert.equal(validateBookingUrl('https://outlook.office.com/bookwithme/user/x').error, undefined);
});

test('Bookings URL on an unexpected host warns but does not block', () => {
  const { url, error, warning } = validateBookingUrl('https://calendly.com/someone/20min');
  assert.equal(error, undefined);
  assert.ok(url);
  assert.match(warning, /does not look like a Microsoft Bookings host/);
});

/* ------------------------------------------------------------------ */
/* escaping and failure modes                                          */
/* ------------------------------------------------------------------ */

test('substituted values are HTML-escaped', () => {
  assert.equal(escapeHtml('a<b>&"\''), 'a&lt;b&gt;&amp;&quot;&#39;');

  const s = settings({ companyName: 'Ben & Jerry <Ltd>' });
  const out = renderTemplate(s.templates['reject-direct'], CANDIDATE, s);
  assert.ok(out.bodyHtml.includes('Ben &amp; Jerry &lt;Ltd&gt;'));
  assert.ok(!out.bodyHtml.includes('<Ltd>'), 'raw markup from a config value reached the body');
});

test('a candidate address containing markup cannot break the body', () => {
  const s = settings();
  const template = { ...s.templates['reject-direct'], bodyHtml: '<p>{{CANDIDATE_EMAIL}}</p>' };
  const out = renderTemplate(template, { email: 'x<script>@example.com' }, s);
  assert.ok(!out.bodyHtml.includes('<script>'));
  assert.ok(out.bodyHtml.includes('&lt;script&gt;'));
});

test('an unknown variable fails rendering instead of shipping a literal', () => {
  const s = settings();
  const template = { ...s.templates['reject-direct'], bodyHtml: '<p>Hello {{NOT_A_THING}}</p>' };
  const out = renderTemplate(template, CANDIDATE, s);
  assert.equal(out.ok, false);
  assert.match(out.errors.join(' '), /\{\{NOT_A_THING\}\} is not a known template variable/);
});

test('a known but unconfigured variable fails rendering', () => {
  const s = settings({ jobTitle: '' });
  const out = renderTemplate(s.templates['reject-direct'], CANDIDATE, s);
  assert.equal(out.ok, false);
  assert.match(out.errors.join(' '), /\{\{JOB_TITLE\}\} has no value configured/);
});

test('empty subject or body fails rendering', () => {
  const s = settings();
  const noSubject = renderTemplate({ ...s.templates['reject-direct'], subject: '  ' }, CANDIDATE, s);
  assert.equal(noSubject.ok, false);
  assert.match(noSubject.errors.join(' '), /Subject is empty/);

  const noBody = renderTemplate({ ...s.templates['reject-direct'], bodyHtml: '<p></p>' }, CANDIDATE, s);
  assert.equal(noBody.ok, false);
  assert.match(noBody.errors.join(' '), /Body is empty/);
});

/* ------------------------------------------------------------------ */
/* greeting handling                                                   */
/* ------------------------------------------------------------------ */

test('greetingMode none renders {{GREETING}} away without erroring', () => {
  const s = settings({ greetingMode: 'none' });
  const template = { ...s.templates['reject-direct'], bodyHtml: '<p>{{GREETING}}</p><p>Body.</p>' };
  const out = renderTemplate(template, CANDIDATE, s);
  assert.equal(out.ok, true, out.errors.join('; '));
  assert.ok(!out.bodyHtml.includes('{{GREETING}}'));
  assert.ok(!/hi\b/i.test(out.bodyText));
});

test('greetingMode name-or-fallback uses the fallback when no name is known', () => {
  const s = settings({ greetingMode: 'name-or-fallback' });
  const template = { ...s.templates['reject-direct'], bodyHtml: '<p>{{GREETING}}</p><p>Body.</p>' };
  const out = renderTemplate(template, { email: 'a@b.com' }, s);
  assert.match(out.bodyText, /^Hi there,/);
});

test('greetingMode name-or-fallback uses the name when Outlook supplied one', () => {
  const s = settings({ greetingMode: 'name-or-fallback' });
  const template = { ...s.templates['reject-direct'], bodyHtml: '<p>{{GREETING}}</p><p>Body.</p>' };
  const out = renderTemplate(template, { email: 'a@b.com', name: 'Jane Roe' }, s);
  assert.match(out.bodyText, /^Hi Jane Roe,/);
});

test('a name is never derived from the email address', () => {
  const s = settings({ greetingMode: 'name-or-fallback' });
  const template = { ...s.templates['reject-direct'], bodyHtml: '<p>{{GREETING}}</p><p>Body.</p>' };
  const out = renderTemplate(template, { email: 'john.doe@gmail.com' }, s);
  assert.ok(!/john/i.test(out.bodyText), 'a name was inferred from the address');
});

/* ------------------------------------------------------------------ */
/* misc                                                                */
/* ------------------------------------------------------------------ */

test('all five actions exist; the two rejections render cleanly out of the box, the rest once configured', () => {
  assert.deepEqual(ACTION_ORDER, [
    'reject-direct',
    'reject-in-process',
    'shortlist-ravilochan',
    'shortlist-abhi',
    'interview-ravilochan',
  ]);
  for (const id of ACTION_ORDER) {
    const out = DEFAULT_SETTINGS.templates[id].requiresBookingUrl
      ? renderWith(id, { bookingUrl: 'https://outlook.office.com/bookwithme/user/fixture' })
      : render(id);
    assert.equal(out.ok, true, `${id}: ${out.errors.join('; ')}`);
  }
});

test('a successful render never leaves a literal {{VARIABLE}} behind', () => {
  for (const id of ACTION_ORDER) {
    for (const mode of ['none', 'name-or-fallback']) {
      const out = render(id, { greetingMode: mode });
      if (!out.ok) continue;
      assert.ok(!out.bodyHtml.includes('{{'), `${id}/${mode} left a placeholder in the body`);
      assert.ok(!out.subject.includes('{{'), `${id}/${mode} left a placeholder in the subject`);
    }
  }
});

test('a suppressed greeting leaves no blank paragraph behind', () => {
  const s = settings({ greetingMode: 'none' });
  const template = { ...s.templates['reject-direct'], bodyHtml: '<p>{{GREETING}}</p><p>Body.</p>' };
  const out = renderTemplate(template, CANDIDATE, s);
  assert.equal(out.bodyHtml, '<p>Body.</p>');
  assert.equal(out.bodyText, 'Body.');
});

test('htmlToText turns paragraphs and breaks into newlines', () => {
  assert.equal(htmlToText('<p>One</p><p>Two<br />Three</p>'), 'One\n\nTwo\nThree');
  assert.equal(htmlToText('<p>&amp; &lt;b&gt;</p>'), '& <b>');
});
