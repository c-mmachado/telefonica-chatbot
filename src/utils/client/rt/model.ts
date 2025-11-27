import z from "zod";

import { Required } from "../../misc";

// export type HyperlinkEntity = Partial<{
//     id: string;
//     name: string;

//     type: HyperlinkType;
//     ref: HyperlinkRef;

//     _url: string;

//     from: string;
//     to: string;
//     label: string;
//     update: string;
// }>;

// export type TypedHyperlinkEntity = HyperlinkEntity & Required<HyperlinkEntity, "type" | "_url">;

// export type RefHyperlinkEntity = HyperlinkEntity & Required<HyperlinkEntity, "ref" | "_url">;

export type CustomFieldHyperlink = TypedHyperlinkEntity &
    Required<HyperlinkEntity, "name"> & {
        values?: string[];
    };

export enum HyperlinkRef {
    Self = "self",
    User = "user",
    Queue = "queue",
    Ticket = "ticket",
    CustomField = "customfield",
    CustomFieldValue = "customfieldvalue",
    Create = "create",
    Comment = "comment",
    History = "history",
}

export enum HyperlinkType {
    User = "user",
    Queue = "queue",
    Ticket = "ticket",
    CustomField = "customfield",
    CustomFieldValue = "customfieldvalue",
    Transaction = "transaction", // ? check
}

// Base type schemas
export const hyperlinkEntitySchema = z.object({
    id: z.string().min(1).optional(),
    name: z.string().min(1).optional(),
    type: z.nativeEnum(HyperlinkType).optional(),
    ref: z.nativeEnum(HyperlinkRef).optional(),
    _url: z.string().url().optional(),
    from: z.string().min(1).optional(),
    to: z.string().min(1).optional(),
    label: z.string().min(1).optional(),
    update: z.string().min(1).optional(),
});
export type HyperlinkEntity = z.infer<typeof hyperlinkEntitySchema>;

export const typedHyperlinkSchema = hyperlinkEntitySchema.and(
    z.object({
        type: z.nativeEnum(HyperlinkType),
        _url: z.string().url(),
    })
);
export type TypedHyperlinkEntity = z.infer<typeof typedHyperlinkSchema>;

export const refHyperlinkSchema = hyperlinkEntitySchema.and(
    z.object({
        ref: z.nativeEnum(HyperlinkRef),
        _url: z.string().url(),
    })
);
export type RefHyperlinkEntity = z.infer<typeof refHyperlinkSchema>;

// Generic factory function to create paged schemas
export const createPagedCollectionSchema = <T extends z.ZodTypeAny>(itemSchema: T) =>
    z.object({
        items: z.array(itemSchema),
        page: z.number().min(1),
        per_page: z.number().min(1),
        total: z.number().min(0),
        pages: z.number().min(1),
        count: z.number().min(0),
        next_page: z.string().url().optional(),
        prev_page: z.string().url().optional(),
    });

export const pagedCollectionSchema = createPagedCollectionSchema(typedHyperlinkSchema);
export type PagedCollection<T> = z.infer<ReturnType<typeof createPagedCollectionSchema<z.ZodType<T>>>>;

// RT root endpoint schema
export const rtSchemaConfig = {
    getResponse: z.object({
        Version: z.string().min(1),
    }),
};
export type RTRoot = z.infer<typeof rtSchemaConfig.getResponse>;

// Queues endpoint schema
export const queueSchema = z.object({
    id: z.string().min(1),
    Name: z.string().min(1),
    TicketCustomFields: typedHyperlinkSchema.array(),
    _hyperlinks: refHyperlinkSchema.array(),
});
export type Queue = z.infer<typeof queueSchema>;

export const queuesSchemaConfig = {
    getResponse: createPagedCollectionSchema(queueSchema),
};

export type QueueRef = TypedHyperlinkEntity;

// export interface Queue {
//     id: string;
//     Name: string;
//     TicketCustomFields: CustomFieldHyperlink[];
//     _hyperlinks: RefHyperlinkEntity[];
// }

// Ticket related types
export type TicketRef = TypedHyperlinkEntity;

export interface Ticket {
    id: string;

    Subject: string;
    Type: string;
    Status: string;

    Requestor: string[];

    InitialPriority: number;
    Priority: number;
    FinalPriority: number;

    TimeLeft: number;
    TimeWorked: number;
    TimeEstimated: number;

    Cc: string[];
    AdminCc: string[];

    Started: Date;
    Resolved: Date;
    Starts: Date;
    Due: Date;
    Created: Date;
    LastUpdated: Date;

    Queue: TypedHyperlinkEntity;
    Owner: TypedHyperlinkEntity;
    Creator: TypedHyperlinkEntity;
    LastUpdatedBy: TypedHyperlinkEntity;

    EffectiveId: TypedHyperlinkEntity;

    CustomFields: CustomFieldHyperlink[];

    _hyperlinks: RefHyperlinkEntity[];
}
export type UpdateTicket = string[];
export type CreateTicket = TicketRef;

// Ticket transaction related types
export type TicketTransactionRef = TypedHyperlinkEntity;

export interface TicketTransaction {}

// Custom fields related types
export type CustomFieldRef = TypedHyperlinkEntity;

export interface CustomField {
    id: string;
    Name: string;
    Description: string;
    Values: string[];
    Type: "Select" | "Freeform" | string;
    Disabled: "0" | "1";
    MaxValues: number;
    Pattern: string;
    EntryHint?: string;
    BasedOn?: TypedHyperlinkEntity;
    Dependents?: CustomField[];
    _hyperlinks: RefHyperlinkEntity[];
}

// Custom fields' values related types
export type CustomFieldValueRef = TypedHyperlinkEntity;

export interface CustomFieldValue {
    id: string;
    Name: string;
    Description: string;
    Category: string;
    _hyperlinks: RefHyperlinkEntity[];
}

// User related types
export type UserRef = TypedHyperlinkEntity;

export interface User {
    id: string;
    Name: string;
    Email: string;
    RealName: string;
    Privileged: "0" | "1";
    _hyperlinks: RefHyperlinkEntity[];
}
