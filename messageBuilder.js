import fs from 'node:fs';
import { contentType as getContentType } from 'mime-types';
export default class MsGraphMailMessageBuilder {
    #attachments = [];
    #bccRecipients = [];
    #body = {
        content: '',
        contentType: 'text'
    };
    #ccRecipients = [];
    #subject = '';
    #toRecipients = [];
    addAttachmentFromBytes(name, contentBytes, contentType) {
        let attachmentContentType = contentType;
        attachmentContentType ??= getContentType(name);
        if (attachmentContentType === false) {
            attachmentContentType = 'application/octet-stream';
        }
        this.#attachments.push({
            '@odata.type': '#microsoft.graph.fileAttachment',
            name,
            contentBytes,
            contentType: attachmentContentType
        });
        return this;
    }
    addAttachmentFromFileBuffer(name, file, contentType) {
        const contentBytes = file.toString('base64');
        return this.addAttachmentFromBytes(name, contentBytes, contentType);
    }
    addAttachmentFromFilePath(name, filePath, contentType) {
        // eslint-disable-next-line security/detect-non-literal-fs-filename
        const contentBytes = fs.readFileSync(filePath).toString('base64');
        return this.addAttachmentFromBytes(name, contentBytes, contentType);
    }
    addBccRecipient(emailAddress) {
        if (typeof emailAddress === 'string') {
            this.#bccRecipients.push({ emailAddress: { address: emailAddress } });
        }
        else {
            this.#bccRecipients.push({ emailAddress });
        }
        return this;
    }
    addCcRecipient(emailAddress) {
        if (typeof emailAddress === 'string') {
            this.#ccRecipients.push({ emailAddress: { address: emailAddress } });
        }
        else {
            this.#ccRecipients.push({ emailAddress });
        }
        return this;
    }
    addToRecipient(emailAddress) {
        if (typeof emailAddress === 'string') {
            this.#toRecipients.push({ emailAddress: { address: emailAddress } });
        }
        else {
            this.#toRecipients.push({ emailAddress });
        }
        return this;
    }
    appendToBody(content, contentType = 'text') {
        if (this.#body.contentType !== contentType) {
            throw new Error(`Cannot append content of type ${contentType} to body with content type ${this.#body.contentType}`);
        }
        this.#body.content += content;
        return this;
    }
    build() {
        const message = {
            subject: this.#subject,
            body: this.#body,
            toRecipients: this.#toRecipients,
            bccRecipients: this.#bccRecipients.length > 0 ? this.#bccRecipients : undefined,
            ccRecipients: this.#ccRecipients.length > 0 ? this.#ccRecipients : undefined,
            attachments: this.#attachments.length > 0 ? this.#attachments : undefined
        };
        return message;
    }
    withBody(content, contentType = 'text') {
        this.#body = {
            content,
            contentType
        };
        return this;
    }
    withSubject(subject) {
        this.#subject = subject;
        return this;
    }
}
