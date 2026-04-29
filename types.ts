import type { Filter } from 'odata-query'

export interface MsGraphMailApiOptions<
  T = MsGraphMailAttachment | MsGraphMailFolder | MsGraphMailMessage
> {
  /** The number of items to skip in the result set. */
  skip?: number

  /** The maximum number of items to return in the result set. */
  top?: number

  /** The list of fields to include in the result set. */
  select?: Array<keyof T>

  /** The filter to apply to the result set. */
  filter?: Filter<T> | string

  /** The order in which to return the items in the result set. */
  orderBy?: Array<`${keyof T & string} ${'asc' | 'desc'}` | (keyof T & string)>
}

export interface MsGraphMailFolder {
  id: string
  displayName: string
  parentFolderId?: string
  childFolderCount?: number
  unreadItemCount?: number
  totalItemCount?: number
  isHidden?: boolean
}

export interface MsGraphMailMessage {
  id: string

  from?: MsGraphMailRecipient
  replyTo?: MsGraphMailRecipient[]
  sender?: MsGraphMailRecipient

  bccRecipients?: MsGraphMailRecipient[]
  ccRecipients?: MsGraphMailRecipient[]
  toRecipients?: MsGraphMailRecipient[]

  importance?: 'high' | 'low' | 'normal'
  subject?: string

  body?: MsGraphMailMessageBody
  bodyPreview?: string
  uniqueBody?: string

  createdDateTime?: string
  receivedDateTime?: string
  sentDateTime?: string

  changeKey?: string
  conversationId?: string
  conversationIndex?: string
  flag: MsGraphMailMessageFlag
  hasAttachments?: boolean
  inferenceClassification?: 'focused' | 'other'
  internetMessageHeaders?: MsGraphMailMessageHeader[]
  internetMessageId?: string
  isDeliveryReceiptRequested?: boolean
  isDraft?: boolean
  isRead?: boolean
  isReadReceiptRequested?: boolean
  lastModifiedDateTime?: string
  parentFolderId?: string
  webLink?: string
}

export interface MsGraphMailAttachment {
  id: string
  contentType: string
  name: string
  size: number
  isInline: boolean
  lastModifiedDateTime?: string

  contentBytes: string
  contentId: string | null
  contentLocation: string | null
}

export interface MsGraphMailRecipient {
  emailAddress: MsGraphMailEmailAddress
}

export interface MsGraphMailEmailAddress {
  name?: string
  address: string
}

export interface MsGraphMailMessageBody {
  contentType: 'html' | 'text'
  content: string
}

export interface MsGraphMailMessageFlag {
  flagStatus: 'complete' | 'flagged' | 'notFlagged'
  completedDateTime?: MsGraphMailDateTimeTimeZone
  dueDateTime?: MsGraphMailDateTimeTimeZone
  startDateTime?: MsGraphMailDateTimeTimeZone
}

export interface MsGraphMailMessageHeader {
  name: string
  value: string
}

export interface MsGraphMailDateTimeTimeZone {
  dateTime: string
  timeZone: string
}

export interface MsGraphMailSendableMessage {
  subject: string

  body: MsGraphMailMessageBody

  toRecipients: MsGraphMailRecipient[]

  bccRecipients?: MsGraphMailRecipient[]
  ccRecipients?: MsGraphMailRecipient[]

  attachments?: MsGraphMailSendableAttachment[]
}

export interface MsGraphMailSendableAttachment {
  '@odata.type': '#microsoft.graph.fileAttachment'
  name: string

  contentBytes: string
  contentType?: string
}
