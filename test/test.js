/* eslint-disable no-console */
import assert from 'node:assert';
import fs from 'node:fs/promises';
import path from 'node:path';
import { describe, it } from 'node:test';
import MsGraphMail, { MsGraphMailMessageBuilder, wellKnownFolderNames } from '../index.js';
import { config, toEmailAddress } from './config.js';
await describe('MsGraphMailApi', async () => {
    await it.skip('should list mail folders', async () => {
        const api = new MsGraphMail(config);
        const mailFolders = await api.listMailFolders();
        console.log(mailFolders);
        assert.ok(Array.isArray(mailFolders), 'Expected mailFolders to be an array');
        assert.ok(mailFolders.length > 0, 'Expected mailFolders to contain at least one folder');
    });
    await it.skip('should get inbox folder', async () => {
        const api = new MsGraphMail(config);
        const inboxFolder = await api.getInboxFolder({
            select: ['displayName', 'id']
        });
        console.log(inboxFolder);
        assert.ok(inboxFolder, 'Expected inboxFolder to be defined');
        assert.strictEqual(inboxFolder.displayName.toLowerCase(), 'inbox', 'Expected inboxFolder displayName to include "Inbox"');
    });
    await it('should list messages in inbox folder', async () => {
        const api = new MsGraphMail(config);
        const messages = await api.listMessages(wellKnownFolderNames.Inbox, {
            select: ['id', 'subject', 'receivedDateTime', 'from', 'body'],
            orderBy: ['receivedDateTime desc'],
            top: 5
        });
        console.log(messages);
        assert.ok(Array.isArray(messages), 'Expected messages to be an array');
    });
    await it.skip('should list messages with attachments in inbox folder', async () => {
        const api = new MsGraphMail(config);
        const messages = await api.listMessages(wellKnownFolderNames.Inbox, {
            filter: {
                hasAttachments: true
            }
        });
        console.log(messages);
        assert.ok(Array.isArray(messages), 'Expected messages to be an array');
        assert.ok(messages.every((message) => message.hasAttachments === true), 'Expected all messages to have attachments');
    });
    await it.skip('should list attachments for a message in inbox folder', async () => {
        const api = new MsGraphMail(config);
        const messages = await api.listMessages(wellKnownFolderNames.Inbox, {
            filter: {
                hasAttachments: true
            },
            orderBy: ['hasAttachments', 'receivedDateTime desc']
        });
        if (messages.length === 0) {
            console.log('No messages with attachments found in inbox folder to list');
            return;
        }
        const messageWithAttachments = messages.at(0);
        if (messageWithAttachments === undefined) {
            console.log('No messages with attachments found in inbox folder to list');
            return;
        }
        const attachments = await api.listMessageAttachments(messageWithAttachments.id);
        console.log(attachments);
        assert.ok(Array.isArray(attachments), 'Expected attachments to be an array');
        for (const attachment of attachments) {
            await fs.writeFile(path.join('test', 'data', attachment.name), Buffer.from(attachment.contentBytes, 'base64'));
            console.log(`Saved attachment "${attachment.name}" to test/data directory`);
        }
    });
    await it('should move a message to the archive folder', async () => {
        const api = new MsGraphMail(config);
        const messages = await api.listMessages(wellKnownFolderNames.Inbox, {
            orderBy: ['receivedDateTime'],
            top: 1
        });
        if (messages.length === 0) {
            console.log('No messages found in inbox folder to move');
            return;
        }
        const messageToMove = messages[0];
        const movedMessage = await api.archiveMessage(messageToMove.id);
        console.log(movedMessage);
        const archiveFolder = await api.getArchiveFolder({
            select: ['id']
        });
        assert.strictEqual(movedMessage.parentFolderId, archiveFolder.id, 'Expected moved message parentFolderId to match archive folder ID');
    });
    await it('should send a message', async () => {
        const api = new MsGraphMail(config);
        const message = new MsGraphMailMessageBuilder()
            .withSubject('Test Message from node-ms-graph-mail')
            .withBody('<p><b>Test Message Body</b></p>', 'html')
            .appendToBody('<p>This message was sent using the node-ms-graph-mail package.</p>', 'html')
            .addToRecipient(toEmailAddress)
            .addAttachmentFromFilePath('logo.png', path.join('test', 'data', 'logo.png'))
            .build();
        await api.sendMessage(message);
        console.log('Message sent successfully');
    });
});
