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
  | 'shortlist-abhi';

/** Order in which the buttons appear in the recruiter panel. */
export const ACTION_ORDER: ActionId[] = [
  'reject-direct',
  'reject-in-process',
  'shortlist-ravilochan',
  'shortlist-abhi',
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
 * Screening run by the account holder. First person, and framed so the 20 minutes
 * goes on technical substance rather than a resume walkthrough — that framing is
 * the whole point of the round.
 */
const SHORTLIST_SELF_BODY = [
  '<p>Thank you for your interest in the {{JOB_TITLE}} position at {{COMPANY}}.</p>',
  GAP,
  '<p>I’ve reviewed your application and would like to move to the next step: a 20-minute technical screening with me over Microsoft Teams.</p>',
  GAP,
  '<p>Rather than going through your resume line by line, I’d like to spend the time on the technical work you’ve actually done:</p>',
  '<ul>\n<li>projects you’ve built, and the decisions you made along the way</li>\n<li>your programming and engineering experience</li>\n<li>your work with AI and full-stack development</li>\n<li>what you want to be working on next, and what drew you to {{COMPANY}}</li>\n</ul>',
  GAP,
  '<p>There’s no coding exercise or written assessment. The only thing worth doing beforehand is picking one or two projects you can talk about in real depth.</p>',
  GAP,
  '<p>Please use the link below to choose a time:</p>\n<p>{{BOOKING_URL}}</p>',
  GAP,
  '<p>Slots can be booked up to one hour before the start time.</p>',
  GAP,
  '<p>Looking forward to hearing how you’ve been building things.</p>',
  SIGNOFF,
].join('\n');

/**
 * Screening delegated to another engineer. The mail still comes from the account
 * holder, so the body names the interviewer and hands over the point of contact.
 */
const SHORTLIST_DELEGATED_BODY = [
  '<p>Thank you for your interest in the {{JOB_TITLE}} position at {{COMPANY}}.</p>',
  GAP,
  '<p>We’ve reviewed your application and would like to move you forward to a 20-minute technical screening with {{INTERVIEWER}}, {{INTERVIEWER_TITLE}} on our engineering team, over Microsoft Teams.</p>',
  GAP,
  '<p>Rather than going through your resume line by line, the conversation will focus on the technical work you’ve actually done:</p>',
  '<ul>\n<li>projects you’ve built, and the decisions you made along the way</li>\n<li>your programming and engineering experience</li>\n<li>your work with AI and full-stack development</li>\n<li>what you want to be working on next, and what drew you to {{COMPANY}}</li>\n</ul>',
  GAP,
  '<p>There’s no coding exercise or written assessment. The only thing worth doing beforehand is picking one or two projects you can talk about in real depth.</p>',
  GAP,
  '<p>Please use the link below to choose a time:</p>\n<p>{{BOOKING_URL}}</p>',
  GAP,
  '<p>Slots can be booked up to one hour before the start time. {{INTERVIEWER}} will be your point of contact for this stage.</p>',
  GAP,
  '<p>Looking forward to hearing how you’ve been building things.</p>',
  SIGNOFF,
].join('\n');

/** Provided by the user. Meant to be sent to candidates. */
const BOOKING_URL_RAVILOCHAN =
  'https://outlook.office.com/bookwithme/user/3f6501d8044d4995ab96268a52c7c6c2@cloudsecurityweb.com/meetingtype/UpQ1JGbg8EuR65lHA9HyMQ2?bookingcode=a9a74792-b0ca-4fec-9922-dbc1cd2c49bd&anonymous&ismsaljsauthenabled&ep=mlink';

const BOOKING_URL_ABHI =
  'https://outlook.office.com/bookwithme/user/ccf48c139cdf40f9b9575ae93ed3ed3f@cloudsecurityweb.com/meetingtype/K6leUTtTw0q22SGhvAvGoA2?anonymous&ismsaljsauthenabled&ep=mcard';

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
      bookingUrl: BOOKING_URL_RAVILOCHAN,
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
      bookingUrl: BOOKING_URL_ABHI,
      interviewerName: 'Abhi',
      interviewerTitle: 'Software Engineer',
      tone: 'positive',
      isPlaceholder: false,
    },
  },

  attemptSourceDeletion: true,
  requireDeletionConfirmation: true,
  debugLogging: false,
};
