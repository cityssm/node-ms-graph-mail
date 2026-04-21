export declare const wellKnownFolderNames: {
    readonly Archive: "Archive";
    readonly Inbox: "Inbox";
    readonly Outbox: "Outbox";
    readonly SentItems: "SentItems";
};
export type WellKnownFolderName = (typeof wellKnownFolderNames)[keyof typeof wellKnownFolderNames];
