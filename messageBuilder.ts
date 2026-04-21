import fs from 'node:fs'

import { contentType as getContentType } from 'mime-types'

import type {
  MsGraphMailEmailAddress,
  MsGraphMailMessageBody,
  MsGraphMailRecipient,
  MsGraphMailSendableAttachment,
  MsGraphMailSendableMessage
} from './types.js'

export default class MsGraphMailMessageBuilder {
  readonly #attachments: MsGraphMailSendableAttachment[] = []

  readonly #bccRecipients: MsGraphMailRecipient[] = []

  #body: MsGraphMailMessageBody = {
    content: '',
    contentType: 'text'
  }

  readonly #ccRecipients: MsGraphMailRecipient[] = []

  #subject = ''

  readonly #toRecipients: MsGraphMailRecipient[] = []

  addAttachmentFromBytes(
    name: string,
    contentBytes: string,
    contentType?: string
  ): this {
    let attachmentContentType: boolean | string | undefined = contentType

    attachmentContentType ??= getContentType(name)

    if (attachmentContentType === false) {
      attachmentContentType = 'application/octet-stream'
    }

    this.#attachments.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name,

      contentBytes,
      contentType: attachmentContentType
    })
    return this
  }

  addAttachmentFromFileBuffer(
    name: string,
    file: Buffer,
    contentType?: string
  ): this {
    const contentBytes = file.toString('base64')
    return this.addAttachmentFromBytes(name, contentBytes, contentType)
  }

  addAttachmentFromFilePath(
    name: string,
    filePath: string,
    contentType?: string
  ): this {
    // eslint-disable-next-line security/detect-non-literal-fs-filename
    const contentBytes = fs.readFileSync(filePath).toString('base64')
    return this.addAttachmentFromBytes(name, contentBytes, contentType)
  }

  addBccRecipient(emailAddress: MsGraphMailEmailAddress | string): this {
    if (typeof emailAddress === 'string') {
      this.#bccRecipients.push({ emailAddress: { address: emailAddress } })
    } else {
      this.#bccRecipients.push({ emailAddress })
    }
    return this
  }

  addCcRecipient(emailAddress: MsGraphMailEmailAddress | string): this {
    if (typeof emailAddress === 'string') {
      this.#ccRecipients.push({ emailAddress: { address: emailAddress } })
    } else {
      this.#ccRecipients.push({ emailAddress })
    }
    return this
  }

  addToRecipient(emailAddress: MsGraphMailEmailAddress | string): this {
    if (typeof emailAddress === 'string') {
      this.#toRecipients.push({ emailAddress: { address: emailAddress } })
    } else {
      this.#toRecipients.push({ emailAddress })
    }
    return this
  }

  appendToBody(content: string, contentType: 'html' | 'text' = 'text'): this {
    if (this.#body.contentType !== contentType) {
      throw new Error(
        `Cannot append content of type ${contentType} to body with content type ${this.#body.contentType}`
      )
    }

    this.#body.content += content

    return this
  }

  build(): MsGraphMailSendableMessage {
    const message: MsGraphMailSendableMessage = {
      subject: this.#subject,

      body: this.#body,

      toRecipients: this.#toRecipients,

      bccRecipients:
        this.#bccRecipients.length > 0 ? this.#bccRecipients : undefined,

      ccRecipients:
        this.#ccRecipients.length > 0 ? this.#ccRecipients : undefined,

      attachments: this.#attachments.length > 0 ? this.#attachments : undefined
    }

    return message
  }

  withBody(content: string, contentType: 'html' | 'text' = 'text'): this {
    this.#body = {
      content,
      contentType
    }

    return this
  }

  withSubject(subject: string): this {
    this.#subject = subject
    return this
  }
}
