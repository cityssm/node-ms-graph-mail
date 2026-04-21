import type { MsGraphMailEmailAddress, MsGraphMailSendableMessage } from './types.js';
export default class MsGraphMailMessageBuilder {
    #private;
    addAttachmentFromBytes(name: string, contentBytes: string, contentType?: string): this;
    addAttachmentFromFileBuffer(name: string, file: Buffer, contentType?: string): this;
    addAttachmentFromFilePath(name: string, filePath: string, contentType?: string): this;
    addBccRecipient(emailAddress: MsGraphMailEmailAddress | string): this;
    addCcRecipient(emailAddress: MsGraphMailEmailAddress | string): this;
    addToRecipient(emailAddress: MsGraphMailEmailAddress | string): this;
    appendToBody(content: string, contentType?: 'html' | 'text'): this;
    build(): MsGraphMailSendableMessage;
    withBody(content: string, contentType?: 'html' | 'text'): this;
    withSubject(subject: string): this;
}
