/* eslint-disable @typescript-eslint/no-unsafe-assignment */
/* eslint-disable @typescript-eslint/no-unsafe-type-assertion */
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import OdataQuery from 'odata-query';
import { wellKnownFolderNames } from './helpers.js';
const buildQuery = OdataQuery;
/**
 * A class for interacting with the Microsoft Graph Mail API,
 * providing methods to manage mail folders and messages for the authenticated user
 * or a specified target user. This class handles authentication using Azure AD
 * and provides convenient methods for common mail operations such as listing folders,
 * retrieving specific folders, listing messages, and moving messages between folders.
 */
export default class MsGraphMailApi {
    #apiUrlRoot;
    #client;
    #clientCredential;
    /**
     * Creates an instance of MsGraphMailApi.
     * @param config - The configuration object for the MsGraphMailApi instance.
     * @param config.tenantId - The tenant ID for Azure AD authentication.
     * @param config.clientId - The client ID for Azure AD authentication.
     * @param config.clientSecret - The client secret for Azure AD authentication.
     * @param config.targetUser - (Optional) The target user for email operations. If not provided, the API will operate on the authenticated user's mailbox.
     * @throws {CredentialUnavailableError} Will throw an error if the provided configuration is invalid or if authentication fails.
     */
    constructor(config) {
        this.#clientCredential = new ClientSecretCredential(config.tenantId, config.clientId, config.clientSecret);
        this.#client = Client.initWithMiddleware({
            authProvider: {
                getAccessToken: async () => {
                    const token = await this.#clientCredential.getToken([
                        'https://graph.microsoft.com/.default'
                    ]);
                    return token.token;
                }
            }
        });
        this.#apiUrlRoot =
            config.targetUser === undefined ? '/me' : `/users/${config.targetUser}`;
    }
    /**
     * Archives a message by moving it to the Archive folder. If the Archive folder does not exist, an error will be thrown.
     * @param messageId - The ID of the message to archive.
     * @returns A promise that resolves to the archived message object after it has been moved to the Archive folder.
     * @throws {Error} Will throw an error if the Archive folder is not found or if the API request fails.
     */
    async archiveMessage(messageId) {
        return await this.moveMessage(messageId, wellKnownFolderNames.Archive);
    }
    /**
     * Retrieves the Archive folder for the authenticated user or the specified target user.
     * If the Archive folder does not exist, an error will be thrown.
     * @param options - (Optional) An object containing options for the API request, such as which fields to select. The 'filter' option is not allowed as its internally set to filter by display name.
     * @returns A promise that resolves to the Archive folder object.
     * @throws {Error} Will throw an error if the Archive folder is not found or if the API request fails.
     */
    async getArchiveFolder(options) {
        const folder = await this.getMailFolderByDisplayName(wellKnownFolderNames.Archive, options);
        if (folder === undefined) {
            throw new Error('Archive folder not found');
        }
        return folder;
    }
    /**
     * Retrieves the Inbox folder for the authenticated user or the specified target user.
     * If the Inbox folder does not exist, an error will be thrown.
     * @param options - (Optional) An object containing options for the API request, such as which fields to select. The 'filter' option is not allowed as its internally set to filter by display name.
     * @returns A promise that resolves to the Inbox folder object.
     * @throws {Error} Will throw an error if the Inbox folder is not found or if the API request fails.
     */
    async getInboxFolder(options) {
        const folder = await this.getMailFolderByDisplayName(wellKnownFolderNames.Inbox, options);
        if (folder === undefined) {
            throw new Error('Inbox folder not found');
        }
        return folder;
    }
    /**
     * Retrieves a mail folder by its display name.
     * @param displayName - The display name of the mail folder to retrieve.
     * @param options - (Optional) An object containing options for the API request, such as which fields to select. The 'filter' option is not allowed as its internally set to filter by display name.
     * @returns A promise that resolves to the mail folder object if found, or undefined if no matching folder is found.
     * @throws {Error} Will throw an error if the API request fails or if the response is not in the expected format.
     */
    async getMailFolderByDisplayName(displayName, options) {
        const folderOptions = {
            filter: {
                displayName: { contains: displayName }
            },
            top: 1,
            ...options
        };
        if (folderOptions.select !== undefined &&
            !folderOptions.select.includes('displayName')) {
            folderOptions.select.push('displayName');
        }
        const mailFolders = await this.listMailFolders(folderOptions);
        const folder = mailFolders.find((folder) => folder.displayName.toLowerCase().includes(displayName.toLowerCase()));
        return folder;
    }
    /**
     * Retrieves the Outbox folder for the authenticated user or the specified target user.
     * If the Outbox folder does not exist, an error will be thrown.
     * @param options - (Optional) An object containing options for the API request, such as which fields to select. The 'filter' option is not allowed as its internally set to filter by display name.
     * @returns A promise that resolves to the Outbox folder object.
     * @throws {Error} Will throw an error if the Outbox folder is not found or if the API request fails.
     */
    async getOutboxFolder(options) {
        const folder = await this.getMailFolderByDisplayName(wellKnownFolderNames.Outbox, options);
        if (folder === undefined) {
            throw new Error('Outbox folder not found');
        }
        return folder;
    }
    /**
     * Retrieves the Sent Items folder for the authenticated user or the specified target user.
     * If the Sent Items folder does not exist, an error will be thrown.
     * @param options - (Optional) An object containing options for the API request, such as which fields to select. The 'filter' option is not allowed as its internally set to filter by display name.
     * @returns A promise that resolves to the Sent Items folder object.
     * @throws {Error} Will throw an error if the Sent Items folder is not found or if the API request fails.
     */
    async getSentItemsFolder(options) {
        const folder = await this.getMailFolderByDisplayName(wellKnownFolderNames.SentItems, options);
        if (folder === undefined) {
            throw new Error('Sent Items folder not found');
        }
        return folder;
    }
    /**
     * Retrieves a list of mail folders for the authenticated user or the specified target user.
     * @param options - (Optional) An object containing options for the API request, such as which fields to select.
     * @returns A promise that resolves to an array of mail folder objects.
     * @see {@link https://learn.microsoft.com/en-us/graph/api/user-list-mailfolders?view=graph-rest-1.0&tabs=http}
     * @throws {Error} Will throw an error if the API request fails or if the response is not in the expected format.
     */
    async listMailFolders(options) {
        return (await this.#callGetApi('/mailFolders', options));
    }
    /**
     * Retrieves a list of attachments for a specific message.
     * @param messageId - The ID of the message for which to retrieve attachments.
     * @param options - (Optional) An object containing options for the API request, such as which fields to select.
     * @returns A promise that resolves to an array of mail attachment objects.
     * @see {@link https://learn.microsoft.com/en-us/graph/api/message-list-attachments?view=graph-rest-1.0&tabs=http}
     * @throws {Error} Will throw an error if the API request fails or if the response is not in the expected format.
     */
    async listMessageAttachments(messageId, options) {
        return (await this.#callGetApi(`/messages/${messageId}/attachments`, options));
    }
    /**
     * Retrieves a list of messages in a specified mail folder for the authenticated user or the specified target user.
     * @param folderIdOrWellKnownFolderName - The ID of the mail folder or a well-known folder name (e.g., 'Inbox', 'Archive') from which to list messages.
     * @param options - (Optional) An object containing options for the API request, such as which fields to select, how to filter the messages, how to order the results, and pagination options like 'skip' and 'top'.
     * @returns A promise that resolves to an array of mail message objects.
     * @see {@link https://learn.microsoft.com/en-us/graph/api/mailfolder-list-messages?view=graph-rest-1.0&tabs=http}
     * @throws {Error} Will throw an error if the API request fails or if the response is not in the expected format.
     */
    async listMessages(folderIdOrWellKnownFolderName, options) {
        return (await this.#callGetApi(`/mailFolders/${folderIdOrWellKnownFolderName}/messages`, options));
    }
    /**
     * Moves a message to a specified mail folder.
     * @param messageId - The ID of the message to move.
     * @param destinationFolderIdOrWellKnownFolderName - The ID of the destination mail folder or a well-known folder name (e.g., 'Inbox', 'Archive').
     * @returns A promise that resolves to the moved mail message object.
     * @see {@link https://learn.microsoft.com/en-us/graph/api/message-move?view=graph-rest-1.0&tabs=http}
     * @throws {Error} Will throw an error if the API request fails or if the response is not in the expected format.
     */
    async moveMessage(messageId, destinationFolderIdOrWellKnownFolderName) {
        const response = await this.#client
            .api(`${this.#apiUrlRoot}/messages/${messageId}/move`)
            .post({ destinationId: destinationFolderIdOrWellKnownFolderName });
        return response;
    }
    /**
     * Sends a message using the Microsoft Graph Mail API.
     * The message should be constructed using the MsGraphMailMessageBuilder to ensure
     * its in the correct format for sending.
     * @param message - The message to be sent, constructed using the MsGraphMailMessageBuilder.
     */
    async sendMessage(message) {
        await this.#client.api(`${this.#apiUrlRoot}/sendMail`).post({ message });
    }
    async #callGetApi(endpoint, options = {}) {
        const api = this.#client.api(`${this.#apiUrlRoot}${endpoint}`);
        if (options.select !== undefined) {
            api.select(options.select.join(','));
        }
        if (options.filter !== undefined) {
            const filterString = typeof options.filter === 'string'
                ? options.filter
                : buildQuery({ filter: options.filter }).slice(9); // Remove leading '$filter='
            // eslint-disable-next-line unicorn/no-array-callback-reference
            api.filter(filterString);
        }
        if (options.orderBy !== undefined) {
            api.orderby(options.orderBy.join(','));
        }
        if (options.skip !== undefined) {
            api.skip(options.skip);
        }
        if (options.top !== undefined) {
            api.top(options.top);
        }
        const response = (await api.get());
        return response.value;
    }
}
export { wellKnownFolderNames } from './helpers.js';
export { default as MsGraphMailMessageBuilder } from './messageBuilder.js';
