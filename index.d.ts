import type { MsGraphMailApiOptions, MsGraphMailAttachment, MsGraphMailFolder, MsGraphMailMessage, MsGraphMailSendableMessage } from './types.js';
export interface MsGraphMailApiConfig {
    tenantId: string;
    clientId: string;
    clientSecret: string;
    targetUser?: string;
}
/**
 * A class for interacting with the Microsoft Graph Mail API,
 * providing methods to manage mail folders and messages for the authenticated user
 * or a specified target user. This class handles authentication using Azure AD
 * and provides convenient methods for common mail operations such as listing folders,
 * retrieving specific folders, listing messages, and moving messages between folders.
 */
export default class MsGraphMailApi {
    #private;
    /**
     * Creates an instance of MsGraphMailApi.
     * @param config - The configuration object for the MsGraphMailApi instance.
     * @param config.tenantId - The tenant ID for Azure AD authentication.
     * @param config.clientId - The client ID for Azure AD authentication.
     * @param config.clientSecret - The client secret for Azure AD authentication.
     * @param config.targetUser - (Optional) The target user for email operations. If not provided, the API will operate on the authenticated user's mailbox.
     * @throws {CredentialUnavailableError} Will throw an error if the provided configuration is invalid or if authentication fails.
     */
    constructor(config: MsGraphMailApiConfig);
    /**
     * Archives a message by moving it to the Archive folder. If the Archive folder does not exist, an error will be thrown.
     * @param messageId - The ID of the message to archive.
     * @returns A promise that resolves to the archived message object after it has been moved to the Archive folder.
     * @throws {Error} Will throw an error if the Archive folder is not found or if the API request fails.
     */
    archiveMessage(messageId: string): Promise<MsGraphMailMessage>;
    /**
     * Retrieves the Archive folder for the authenticated user or the specified target user.
     * If the Archive folder does not exist, an error will be thrown.
     * @param options - (Optional) An object containing options for the API request, such as which fields to select. The 'filter' option is not allowed as its internally set to filter by display name.
     * @returns A promise that resolves to the Archive folder object.
     * @throws {Error} Will throw an error if the Archive folder is not found or if the API request fails.
     */
    getArchiveFolder(options?: Pick<MsGraphMailApiOptions<MsGraphMailFolder>, 'select'>): Promise<MsGraphMailFolder>;
    /**
     * Retrieves the Inbox folder for the authenticated user or the specified target user.
     * If the Inbox folder does not exist, an error will be thrown.
     * @param options - (Optional) An object containing options for the API request, such as which fields to select. The 'filter' option is not allowed as its internally set to filter by display name.
     * @returns A promise that resolves to the Inbox folder object.
     * @throws {Error} Will throw an error if the Inbox folder is not found or if the API request fails.
     */
    getInboxFolder(options?: Pick<MsGraphMailApiOptions<MsGraphMailFolder>, 'select'>): Promise<MsGraphMailFolder>;
    /**
     * Retrieves a mail folder by its display name.
     * @param displayName - The display name of the mail folder to retrieve.
     * @param options - (Optional) An object containing options for the API request, such as which fields to select. The 'filter' option is not allowed as its internally set to filter by display name.
     * @returns A promise that resolves to the mail folder object if found, or undefined if no matching folder is found.
     * @throws {Error} Will throw an error if the API request fails or if the response is not in the expected format.
     */
    getMailFolderByDisplayName(displayName: string, options?: Pick<MsGraphMailApiOptions<MsGraphMailFolder>, 'select'>): Promise<MsGraphMailFolder | undefined>;
    /**
     * Retrieves the Outbox folder for the authenticated user or the specified target user.
     * If the Outbox folder does not exist, an error will be thrown.
     * @param options - (Optional) An object containing options for the API request, such as which fields to select. The 'filter' option is not allowed as its internally set to filter by display name.
     * @returns A promise that resolves to the Outbox folder object.
     * @throws {Error} Will throw an error if the Outbox folder is not found or if the API request fails.
     */
    getOutboxFolder(options?: Pick<MsGraphMailApiOptions<MsGraphMailFolder>, 'select'>): Promise<MsGraphMailFolder>;
    /**
     * Retrieves the Sent Items folder for the authenticated user or the specified target user.
     * If the Sent Items folder does not exist, an error will be thrown.
     * @param options - (Optional) An object containing options for the API request, such as which fields to select. The 'filter' option is not allowed as its internally set to filter by display name.
     * @returns A promise that resolves to the Sent Items folder object.
     * @throws {Error} Will throw an error if the Sent Items folder is not found or if the API request fails.
     */
    getSentItemsFolder(options?: Pick<MsGraphMailApiOptions<MsGraphMailFolder>, 'select'>): Promise<MsGraphMailFolder>;
    /**
     * Retrieves a list of mail folders for the authenticated user or the specified target user.
     * @param options - (Optional) An object containing options for the API request, such as which fields to select.
     * @returns A promise that resolves to an array of mail folder objects.
     * @see {@link https://learn.microsoft.com/en-us/graph/api/user-list-mailfolders?view=graph-rest-1.0&tabs=http}
     * @throws {Error} Will throw an error if the API request fails or if the response is not in the expected format.
     */
    listMailFolders(options?: MsGraphMailApiOptions<MsGraphMailFolder>): Promise<MsGraphMailFolder[]>;
    /**
     * Retrieves a list of attachments for a specific message.
     * @param messageId - The ID of the message for which to retrieve attachments.
     * @param options - (Optional) An object containing options for the API request, such as which fields to select.
     * @returns A promise that resolves to an array of mail attachment objects.
     * @see {@link https://learn.microsoft.com/en-us/graph/api/message-list-attachments?view=graph-rest-1.0&tabs=http}
     * @throws {Error} Will throw an error if the API request fails or if the response is not in the expected format.
     */
    listMessageAttachments(messageId: string, options?: MsGraphMailApiOptions<MsGraphMailAttachment>): Promise<MsGraphMailAttachment[]>;
    /**
     * Retrieves a list of messages in a specified mail folder for the authenticated user or the specified target user.
     * @param folderIdOrWellKnownFolderName - The ID of the mail folder or a well-known folder name (e.g., 'Inbox', 'Archive') from which to list messages.
     * @param options - (Optional) An object containing options for the API request, such as which fields to select, how to filter the messages, how to order the results, and pagination options like 'skip' and 'top'.
     * @returns A promise that resolves to an array of mail message objects.
     * @see {@link https://learn.microsoft.com/en-us/graph/api/mailfolder-list-messages?view=graph-rest-1.0&tabs=http}
     * @throws {Error} Will throw an error if the API request fails or if the response is not in the expected format.
     */
    listMessages(folderIdOrWellKnownFolderName: string, options?: MsGraphMailApiOptions<MsGraphMailMessage>): Promise<MsGraphMailMessage[]>;
    /**
     * Marks a message as read by setting its 'isRead' property to true.
     * @param messageId - The ID of the message to mark as read.
     * @returns A promise that resolves when the operation is complete.
     * @see {@link https://learn.microsoft.com/en-us/graph/api/message-update?view=graph-rest-1.0&tabs=http}
     * @throws {Error} Will throw an error if the API request fails or if the response is not in the expected format.
     */
    markMessageAsRead(messageId: string): Promise<void>;
    /**
     * Moves a message to a specified mail folder.
     * @param messageId - The ID of the message to move.
     * @param destinationFolderIdOrWellKnownFolderName - The ID of the destination mail folder or a well-known folder name (e.g., 'Inbox', 'Archive').
     * @returns A promise that resolves to the moved mail message object.
     * @see {@link https://learn.microsoft.com/en-us/graph/api/message-move?view=graph-rest-1.0&tabs=http}
     * @throws {Error} Will throw an error if the API request fails or if the response is not in the expected format.
     */
    moveMessage(messageId: string, destinationFolderIdOrWellKnownFolderName: string): Promise<MsGraphMailMessage>;
    /**
     * Sends a message using the Microsoft Graph Mail API.
     * The message should be constructed using the MsGraphMailMessageBuilder to ensure
     * its in the correct format for sending.
     * @param message - The message to be sent, constructed using the MsGraphMailMessageBuilder.
     * @returns A promise that resolves when the operation is complete.
     * @see {@link https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http}
     * @throws {Error} Will throw an error if the API request fails or if the response is not in the expected format.
     */
    sendMessage(message: MsGraphMailSendableMessage): Promise<void>;
}
export { wellKnownFolderNames } from './helpers.js';
export { default as MsGraphMailMessageBuilder } from './messageBuilder.js';
export type { MsGraphMailApiOptions, MsGraphMailFolder, MsGraphMailMessage } from './types.js';
