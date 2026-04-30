# Microsoft Graph Mail API

[![DeepSource](https://app.deepsource.com/gh/cityssm/node-ms-graph-mail.svg/?label=active+issues&show_trend=true&token=KkSuPUghHqeLWICUprgijw-o)](https://app.deepsource.com/gh/cityssm/node-ms-graph-mail/)

Wrappers around the Microsoft Graph API
to add response types to the Mail-related queries.

## Installation

```sh
npm install @cityssm/ms-graph-mail
```

## Usage

```javascript
import MsGraphMail, {
  MsGraphMailMessageBuilder,
  wellKnownFolderNames
} from '@cityssm/ms-graph-mail'

const api = new MsGraphMail({
  tenantId: '00000000-0000-0000-0000-00000000000a',
  clientId: '00000000-0000-0000-0000-00000000000b',
  clientSecret: 'abcd...xyz',
  targetUser: 'helpdesk@example.com'
})

const inboxMessages = await api.listMessages(wellKnownFolderNames.Inbox, {
  top: 5,
  select: ['id', 'subject', 'receivedDateTime', 'from', 'body'],
  orderBy: ['receivedDateTime desc']
})

const messageToSend = new MsGraphMailMessageBuilder()
  .withSubject('Ticket #000001')
  .withBody('<p><b>Ticket Received</b></p>', 'html')
  .appendToBody('<p>Your ticket should be worked on shortly.</p>', 'html')
  .addToRecipient('requestor@example.com')
  .build()

await api.sendMessage(messageToSend)
```

## Functions

👍 It is recommended to use this package with Typescript to get usage hints.

### Folder Functions

```typescript
async function listMailFolders(
  options?: MsGraphMailApiOptions<MsGraphMailFolder>
): Promise<MsGraphMailFolder[]> {}

async function getMailFolderByDisplayName(
  displayName: string,
  options?: Pick<MsGraphMailApiOptions<MsGraphMailFolder>, 'select'>
): Promise<MsGraphMailFolder | undefined> {}

async function getArchiveFolder(
  options?: Pick<MsGraphMailApiOptions<MsGraphMailFolder>, 'select'>
): Promise<MsGraphMailFolder> {}

async function getInboxFolder(
  options?: Pick<MsGraphMailApiOptions<MsGraphMailFolder>, 'select'>
): Promise<MsGraphMailFolder> {}

async function getOutboxFolder(
  options?: Pick<MsGraphMailApiOptions<MsGraphMailFolder>, 'select'>
): Promise<MsGraphMailFolder> {}

async function getSentItemsFolder(
  options?: Pick<MsGraphMailApiOptions<MsGraphMailFolder>, 'select'>
): Promise<MsGraphMailFolder> {}
```

### Message Functions

```typescript
async function listMessages(
  folderIdOrWellKnownFolderName: string,
  options?: MsGraphMailApiOptions<MsGraphMailMessage>
): Promise<MsGraphMailMessage[]> {}

async function listMessageAttachments(
  messageId: string,
  options?: MsGraphMailApiOptions<MsGraphMailAttachment>
): Promise<MsGraphMailAttachment[]> {}
```

### Message Move Functions

```typescript
async function archiveMessage(messageId: string): Promise<MsGraphMailMessage> {}

async function moveMessage(
  messageId: string,
  destinationFolderIdOrWellKnownFolderName: string
): Promise<MsGraphMailMessage> {}
```

### Message Send Functions

```typescript
async function sendMessage(
  message: MsGraphMailSendableMessage
): Promise<void> {}
```

### Message Update Functions

```typescript
async function markMessageAsRead(messageId: string): Promise<void> {}
```

## Related Projects

[**ShiftLog**](https://github.com/cityssm/shiftLog/)<br />
A work management system with work order recording, shift activity logging,
and timesheet tracking.
