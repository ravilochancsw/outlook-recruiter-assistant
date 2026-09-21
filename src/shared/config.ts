/**
 * Configuration and email templates.
 *
 * The template copy below is the user's own wording, recorded verbatim from the
 * Outlook templates they already use. Placeholders are marked explicitly — do
 * not invent replacements for them.
 */

/**
 * Shortlisting splits by *who runs the screening*, because that is the only
 * difference between the two positive paths: each routes to a different
 * Microsoft Bookings link. Adding a third interviewer means adding an id here,
 * a template below, and nothing else.
 */
export type ActionId =
  | 'reject-direct'
  | 'reject-in-process'
  | 'shortlist-ravilochan'
  | 'shortlist-abhi'
  | 'interview-ravilochan';

/** Order in which the buttons appear in the recruiter panel. */
export const ACTION_ORDER: ActionId[] = [
  'reject-direct',
  'reject-in-process',
  'shortlist-ravilochan',
  'shortlist-abhi',
  // Appended rather than slotted in next to the screenings: the panel's Ctrl+N
  // shortcuts are positional, so inserting here would silently remap them.
  'interview-ravilochan',
];

export interface EmailTemplate {
  id: ActionId;
  /** Button label in the recruiter panel. Kept short. */
  label: string;
  /** Longer explanation, shown as the button's title/tooltip. */
  description: string;
  subject: string;
  /**
   * Body as HTML. Substituted values are HTML-escaped; {{BOOKING_URL}} becomes
   * an anchor. Outlook's editor is rich text, and the user's own reject template
   * bolds the job title and company, so HTML is the native format here.
   */
  bodyHtml: string;
  /** Filling is refused when true and this template has no valid Bookings URL. */
  requiresBookingUrl: boolean;
  /**
   * Per-template Microsoft Bookings link. Lives on the template rather than
   * globally because the whole point of having two shortlist actions is that
   * they book different people's calendars.
   */
  bookingUrl?: string;
  /** Who runs the screening this template books. Exposed as {{INTERVIEWER}}. */
  interviewerName?: string;
  /** Their role, for the candidate's benefit. Exposed as {{INTERVIEWER_TITLE}}. */
  interviewerTitle?: string;
  /** Visual weight in the panel. Destructive-ish actions read differently. */
  tone: 'negative' | 'positive';
  /** True when the copy has not yet been confirmed by the user. */
  isPlaceholder: boolean;
}

export interface Settings {
  /* identity — used for template variables */
  jobTitle: string;
  companyName: string;
  senderName: string;
  senderTitle: string;

  /* greeting handling */
  greetingMode: 'none' | 'name-or-fallback';
  greetingFallback: string;

  /* templates */
  templates: Record<ActionId, EmailTemplate>;

  /* safety */
  attemptSourceDeletion: boolean;
  requireDeletionConfirmation: boolean;

  /* troubleshooting */
  debugLogging: boolean;
}

const SIGNOFF = '<p>Best regards,<br />{{SENDER_NAME}}<br />{{SENDER_TITLE}}</p>';

/**
 * An intentional blank line between paragraphs. `<p>&nbsp;</p>` survives both the
 * template renderer's empty-paragraph cleanup and Outlook's editor, which is why
 * it is used instead of margin CSS that an email client would strip.
 */
const GAP = '<p>&nbsp;</p>';

/**
 * Verbatim from the user's Outlook template "Update on Your Application".
 * Bold matches the original: job title and company are emphasised.
 */
const REJECT_DIRECT_BODY = [
  '<p>Thank you for your interest in <strong>{{JOB_TITLE}}</strong> position at <strong>{{COMPANY}}</strong> and for taking the time to apply.</p>',
  GAP,
  '<p>After careful consideration, we’ve decided to move forward with other candidates whose skills and experience more closely match our current needs.</p>',
  GAP,
  '<p>We appreciate your time and interest in joining our team, and we wish you all the best in your job search.</p>',
  GAP,
  SIGNOFF,
].join('\n');

/**
 * Verbatim from the user's second template, for candidates rejected after
 * entering the interview process. No bold in the original.
 */
const REJECT_IN_PROCESS_BODY = [
  '<p>Thank you for taking the time to apply for the {{JOB_TITLE}} position at {{COMPANY}}.</p>',
  GAP,
  '<p>We were impressed by your background and experience. After careful consideration, however, we’ve decided to move forward with other candidates whose qualifications more closely align with what we’re looking for in this particular role.</p>',
  GAP,
  '<p>We appreciate the time and effort you put into your application and will keep your application on file. We encourage you to consider future opportunities with us that may be a strong match for your experience and qualifications. Should a suitable opportunity arise, we would be happy to reach out to you.</p>',
  GAP,
  '<p>We wish you continued success in your job search.</p>',
  GAP,
  SIGNOFF,
].join('\n');

/**
 * Shared body of the first-round screening.
 *
 * Deliberate omissions, each one a decision:
 *
 * - **No mention of later rounds.** What comes next is disclosed if the candidate
 *   progresses, or answered live if they ask. Naming it here invites questions
 *   that do not belong in a 20-minute screen.
 * - **No "why do you want to work here".** The email promises not to walk the
 *   resume, and that question produces the most rehearsed answer in interviewing.
 *   It is also not part of the 20 minutes.
 * - **No seniority signalling.** A candidate told the interviewer is busy performs
 *   below their real level, which costs signal in exactly the time being
 *   protected. Brevity and precision carry the weight instead.
 *
 * The three bullets are the whole point: they are the things a candidate cannot
 * rehearse from a job description, and they are what the round actually probes.
 */
const SCREENING_TOPICS = `<ul>
<li>what you built, and the part you worked on yourself</li>
<li>why you built it that way, and what you considered instead</li>
<li>what you would do differently now</li>
</ul>`;

const SCREENING_PREP =
  '<p>There’s no coding exercise and nothing to prepare beyond that.</p>';

const SCREENING_BOOKING = [
  '<p>Please use the link below to choose a time that suits you. Slots can be booked up to one hour before they start.</p>',
  '<p>{{BOOKING_URL}}</p>',
].join('\n');

const SCREENING_CLOSE = '<p>Looking forward to hearing how you’ve been building things.</p>';

/** Run by the account holder, so first person throughout. */
const SHORTLIST_SELF_BODY = [
  '<p>Thank you for your interest in the {{JOB_TITLE}} position at {{COMPANY}}.</p>',
  GAP,
  '<p>I’ve reviewed your application and would like to set up a 20-minute technical screening over Microsoft Teams.</p>',
  GAP,
  '<p>Rather than going through your resume line by line, I’d like to spend the time on what you’ve actually built. Pick one or two projects you know well and be ready to go into detail on:</p>',
  SCREENING_TOPICS,
  GAP,
  SCREENING_PREP,
  GAP,
  SCREENING_BOOKING,
  GAP,
  SCREENING_CLOSE,
  SIGNOFF,
].join('\n');

/**
 * Same round, same agenda, conducted by another engineer. The mail still comes
 * from the account holder, so only the interviewer and the point of contact
 * change.
 */
const SHORTLIST_DELEGATED_BODY = [
  '<p>Thank you for your interest in the {{JOB_TITLE}} position at {{COMPANY}}.</p>',
  GAP,
  '<p>We’ve reviewed your application and would like to set up a 20-minute technical screening over Microsoft Teams with {{INTERVIEWER}}, {{INTERVIEWER_TITLE}} on our engineering team.</p>',
  GAP,
  '<p>Rather than going through your resume line by line, the conversation will focus on what you’ve actually built. Pick one or two projects you know well and be ready to go into detail on:</p>',
  SCREENING_TOPICS,
  GAP,
  SCREENING_PREP,
  GAP,
  SCREENING_BOOKING,
  GAP,
  '<p>{{INTERVIEWER}} will be your point of contact for this conversation.</p>',
  GAP,
  SCREENING_CLOSE,
  SIGNOFF,
].join('\n');

/**
 * Round two: the hour where the decision actually gets made.
 *
 * Two things separate it from the screening copy. It does **not** promise "no
 * coding exercise" — this round asks real technical questions, and a promise the
 * interviewer intends to break would both mislead the candidate and skew the
 * conversation. And the prep instruction is the opposite: for the screening the
 * only ask is to pick a project; here the candidate is told to be ready to open
 * one up end to end, including where it went wrong.
 *
 * It still says nothing about what follows, for the same reason the screening does
 * not: what comes next is disclosed on progression, or answered live if asked.
 */
const INTERVIEW_BODY = [
  '<p>Thank you for taking the time to speak with us.</p>',
  GAP,
  '<p>I’d like to move to a one-hour technical discussion over Microsoft Teams.</p>',
  GAP,
  '<p>This one goes deeper. Expect to spend the time on:</p>',
  `<ul>
<li>a walkthrough of something you’ve built, in as much depth as it goes</li>
<li>how you would design and structure a system, and the trade-offs you would weigh</li>
<li>implementation detail on the parts you know best</li>
</ul>`,
  GAP,
  '<p>Worth doing beforehand: pick one project you can open up and talk through end to end, including the parts that gave you trouble.</p>',
  GAP,
  SCREENING_BOOKING,
  GAP,
  '<p>Looking forward to it.</p>',
  SIGNOFF,
].join('\n');

/**
 * Bookings links are per-installation, not shipped. Each user's real link
 * identifies their own Microsoft 365 mailbox and calendar, so it belongs in
 * this installation's Options (`chrome.storage.local`), never in source —
 * hardcoding one here would land it in version control for good. Until a
 * link is configured, `renderTemplate` fails closed rather than sending a
 * candidate a template with no working link.
 */
const BOOKING_URL_UNSET = undefined;

export const DEFAULT_SETTINGS: Settings = {
  jobTitle: 'AI Full Stack Software Engineer',
  companyName: 'Cloud Security Web',
  senderName: 'Ravilochan',
  senderTitle: 'Head of Engineering',

  greetingMode: 'none',
  greetingFallback: 'there',

  templates: {
    'reject-direct': {
      id: 'reject-direct',
      label: 'Reject',
      description: 'Direct rejection — candidate did not progress past the resume review.',
      subject: 'Update on Your Application',
      bodyHtml: REJECT_DIRECT_BODY,
      requiresBookingUrl: false,
      tone: 'negative',
      isPlaceholder: false,
    },
    'reject-in-process': {
      id: 'reject-in-process',
      label: 'Reject (in process)',
      description:
        'Rejection after the candidate entered the interview process, or was not shortlisted for further rounds.',
      subject: 'Update in Interview Process',
      bodyHtml: REJECT_IN_PROCESS_BODY,
      requiresBookingUrl: false,
      tone: 'negative',
      isPlaceholder: false,
    },
    'shortlist-ravilochan': {
      id: 'shortlist-ravilochan',
      label: 'Shortlist → Ravilochan',
      description: '20-minute technical screening with you, on your calendar.',
      subject: 'Technical Screening – {{JOB_TITLE}}',
      bodyHtml: SHORTLIST_SELF_BODY,
      requiresBookingUrl: true,
      bookingUrl: BOOKING_URL_UNSET,
      interviewerName: 'Ravilochan',
      interviewerTitle: 'Head of Engineering',
      tone: 'positive',
      isPlaceholder: false,
    },
    'shortlist-abhi': {
      id: 'shortlist-abhi',
      label: 'Shortlist → Abhi',
      description: '20-minute technical screening delegated to Abhi, on his calendar.',
      subject: 'Technical Screening – {{JOB_TITLE}}',
      bodyHtml: SHORTLIST_DELEGATED_BODY,
      requiresBookingUrl: true,
      bookingUrl: BOOKING_URL_UNSET,
      interviewerName: 'Abhi',
      interviewerTitle: 'Software Engineer',
      tone: 'positive',
      isPlaceholder: false,
    },
    'interview-ravilochan': {
      id: 'interview-ravilochan',
      label: 'Interview (1h)',
      description:
        'Round two: one-hour deep technical discussion on Ravilochan’s calendar — design, implementation and a project walkthrough.',
      subject: 'Technical Interview – {{JOB_TITLE}}',
      bodyHtml: INTERVIEW_BODY,
      requiresBookingUrl: true,
      bookingUrl: BOOKING_URL_UNSET,
      interviewerName: 'Ravilochan',
      interviewerTitle: 'Head of Engineering',
      tone: 'positive',
      isPlaceholder: false,
    },
  },

  attemptSourceDeletion: true,
  requireDeletionConfirmation: true,
  debugLogging: false,
};
