/** Typed messages between content scripts and the service worker. */

export interface SourceSnapshot {
  /** Outlook's own message id, taken from the reading-pane URL path. */
  messageId: string;
  /** Sender display name / address as rendered in the reading pane. */
  sender: string;
  subject: string;
  /** True when sender + subject both look like a LinkedIn application email. */
  looksLikeApplication: boolean;
  capturedAt: number;
  /** Which reading-pane strategy matched, for troubleshooting. */
  paneVia?: string;
}

export type WorkflowState =
  | 'COMPOSE_DETECTED'
  | 'EMAIL_FILLED'
  | 'SEND_CLICKED'
  | 'SEND_CONFIRMED'
  | 'SEND_ABANDONED'
  | 'DELETE_OFFERED'
  | 'DELETED'
  | 'DELETE_DECLINED'
  | 'DELETE_UNSAFE';

export interface Workflow {
  id: string;
  composeTabId: number;
  /** Window the compose tab lives in, for the fallback source-tab search. */
  composeWindowId?: number;
  /** From chrome.tabs openerTabId — an authoritative link, not a guess. */
  sourceTabId?: number;
  /** How the source tab was identified, for the confirmation prompt. */
  sourceVia?: 'opener' | 'sole-application-tab';
  /** What the source tab had open when the compose tab was created. */
  sourceSnapshot?: SourceSnapshot;
  candidateEmail: string;
  actionLabel: string;
  state: WorkflowState;
  createdAt: number;
  /** When Send was clicked, for the freshness bound on the tab-close path. */
  sendClickedAt?: number;
}

/* ---------------- content → background ---------------- */

export interface RegisterFillMsg {
  type: 'RA_REGISTER_FILL';
  candidateEmail: string;
  actionLabel: string;
}

export interface SendClickedMsg {
  type: 'RA_SEND_CLICKED';
}

export interface DiscardClickedMsg {
  type: 'RA_DISCARD_CLICKED';
}

export interface SendConfirmedMsg {
  type: 'RA_SEND_CONFIRMED';
  /** Which signals fired, for the confirmation UI's audit line. */
  signals: string[];
}

export interface SendAbandonedMsg {
  type: 'RA_SEND_ABANDONED';
  reason: string;
}

export interface SnapshotReplyMsg {
  type: 'RA_SNAPSHOT_REPLY';
  snapshot: SourceSnapshot | null;
}

export interface DeleteResultMsg {
  type: 'RA_DELETE_RESULT';
  workflowId: string;
  ok: boolean;
  detail: string;
}

/* ---------------- background → content ---------------- */

export interface CaptureSnapshotMsg {
  type: 'RA_CAPTURE_SNAPSHOT';
}

export interface OfferDeletionMsg {
  type: 'RA_OFFER_DELETION';
  workflowId: string;
  candidateEmail: string;
  actionLabel: string;
  snapshot: SourceSnapshot;
  requireConfirmation: boolean;
  signals: string[];
}

export interface PingMsg {
  type: 'RA_PING';
}

export interface CleanupStatusMsg {
  type: 'RA_CLEANUP_STATUS';
  outcome: 'offered' | 'blocked';
  reason: string;
}

export interface PongMsg {
  type: 'RA_PONG';
  version: string;
  host: string;
  path: string;
  topFrame: boolean;
  assistantMounted: boolean;
  composeDetected: boolean;
  candidateEmail?: string;
  sourceMessageOpen: boolean;
  /** What a snapshot of this tab would report, so cleanup problems are visible. */
  snapshot?: SourceSnapshot | null;
}

export type ContentToBackground =
  | RegisterFillMsg
  | SendClickedMsg
  | DiscardClickedMsg
  | SendConfirmedMsg
  | SendAbandonedMsg
  | DeleteResultMsg;

export type BackgroundToContent = CaptureSnapshotMsg | OfferDeletionMsg | CleanupStatusMsg;

export type AnyMessage = ContentToBackground | BackgroundToContent | PingMsg;

/** Narrow an unknown message payload. */
export function messageType(message: unknown): string | undefined {
  if (typeof message === 'object' && message !== null && 'type' in message) {
    const type = (message as { type?: unknown }).type;
    return typeof type === 'string' ? type : undefined;
  }
  return undefined;
}
