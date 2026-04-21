export const wellKnownFolderNames = {
  Archive: 'Archive',
  Inbox: 'Inbox',
  Outbox: 'Outbox',
  SentItems: 'SentItems'
} as const

export type WellKnownFolderName = (typeof wellKnownFolderNames)[keyof typeof wellKnownFolderNames]
