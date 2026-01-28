import { z } from "zod";

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

// Base types schemas
export const hyperlinkSchema = z.object({
    id: z.string().min(1).optional(), // TODO: number or string -> string coercion
    name: z.string().min(1).optional(),
    type: z.nativeEnum(HyperlinkType).optional(),
    ref: z.nativeEnum(HyperlinkRef).optional(),
    _url: z.string().url().optional(),
    from: z.string().min(1).optional(),
    to: z.string().min(1).optional(),
    label: z.string().min(1).optional(),
    update: z.string().min(1).optional(),
});
export interface HyperlinkEntity extends z.infer<typeof hyperlinkSchema> {}

export const typedHyperlinkSchema = hyperlinkSchema.and(
    z.object({
        type: z.nativeEnum(HyperlinkType),
        _url: z.string().url(),
    })
);
export interface TypedHyperlinkEntity extends z.infer<typeof typedHyperlinkSchema> {}

export const refHyperlinkSchema = hyperlinkSchema.and(
    z.object({
        ref: z.nativeEnum(HyperlinkRef),
        _url: z.string().url(),
    })
);
export interface RefHyperlinkEntity extends z.infer<typeof refHyperlinkSchema> {}
