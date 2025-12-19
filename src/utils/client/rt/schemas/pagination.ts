import { z } from "zod";

import {
    Client,
    BaseSchemaEndpointConfigurer,
    SchemaEndpointConfigurer,
    SchemaEndpointConfig,
    Callbacks,
} from "../../base";
import { HyperlinkEntity } from "./base";

/**
 * Generic factory function to create paged schemas for RT collections used
 * to allow a strongly typed paged collection response.
 *
 * @param itemSchema Schema for the items in the collection
 * @returns Zod schema for the paged collection
 */
export const createRTPagedCollectionSchema = <T extends z.ZodTypeAny>(itemSchema: T) => {
    return z.object({
        items: z.array(itemSchema),
        page: z.number().min(1),
        per_page: z.number().min(1),
        total: z.number().min(0),
        pages: z.number().min(1),
        count: z.number().min(0),
        next_page: z.string().url().optional(),
        prev_page: z.string().url().optional(),
    });
};

interface PagedCollection<_Item> {
    // Intentionally left empty
}

export interface RTPagedCollection<Item>
    extends PagedCollection<Item>,
        z.infer<ReturnType<typeof createRTPagedCollectionSchema<z.ZodType<Item>>>> {}

export const createRTNavigatablePagedCollectionSchema = <T extends z.ZodTypeAny, Config extends SchemaEndpointConfig>(
    itemSchema: T
): z.ZodType<RTNavigatablePagedCollection<Config, z.infer<T>>> => {
    type Item = z.infer<T>;

    return createRTPagedCollectionSchema(itemSchema) as unknown as z.ZodType<
        RTPagedCollection<Item> & {
            next(): SchemaEndpointConfigurer<Config>;

            prev(): SchemaEndpointConfigurer<Config>;
        }
    >;
};

type InferItemFromConfig<Config extends SchemaEndpointConfig> = Config extends {
    methods: {
        get: {
            response: infer _Item extends z.ZodType<infer ZodItem>;
        };
    };
}
    ? ZodItem
    : Config extends {
          methods: {
              get: {
                  response: infer Item;
              };
          };
      }
    ? Item
    : never;

export interface RTNavigatablePagedCollection<Config extends SchemaEndpointConfig, Item = InferItemFromConfig<Config>>
    extends RTPagedCollection<Item> {
    items: Item[];

    next(): SchemaEndpointConfigurer<Config>;

    prev(): SchemaEndpointConfigurer<Config>;
}

export function createRTNavigatablePagedCollection<
    Config extends SchemaEndpointConfig,
    Item = InferItemFromConfig<Config>
>(
    client: Client,
    config: Config,
    callbacks: Callbacks<Config>,
    page: RTPagedCollection<Item>
): RTNavigatablePagedCollection<Config, Item> {
    return new DefaultRTNavigatablePagedCollection(client, config, callbacks, page);
}

class DefaultRTNavigatablePagedCollection<Config extends SchemaEndpointConfig, Item = InferItemFromConfig<Config>>
    implements RTNavigatablePagedCollection<Config, Item>
{
    constructor(
        private readonly _client: Client,
        private readonly _config: Config,
        private readonly _callbacks: Callbacks<Config>,
        private readonly _page: RTPagedCollection<Item>
    ) {}

    public get items(): Item[] {
        return this._page.items;
    }

    public get page(): number {
        return this._page.page;
    }

    public get per_page(): number {
        return this._page.per_page;
    }

    public get total(): number {
        return this._page.total;
    }

    public get pages(): number {
        return this._page.pages;
    }

    public get count(): number {
        return this._page.count;
    }

    public get next_page(): string | undefined {
        return this._page.next_page;
    }

    public get prev_page(): string | undefined {
        return this._page.prev_page;
    }

    public next(): SchemaEndpointConfigurer<Config> {
        if (!this._client || !this._config) {
            throw new Error("Cannot get next page: No client or config available.");
        }

        const nextPage = this.next_page;
        // TODO: Callbacks should also be propagated to the new instance
        return new (class extends BaseSchemaEndpointConfigurer<Config> implements SchemaEndpointConfigurer<Config> {})(
            this._client,
            {
                ...this._config,
                path: nextPage,
            },
            this._callbacks
        );
    }

    public prev(): SchemaEndpointConfigurer<Config> {
        if (!this._client || !this._config) {
            throw new Error("Cannot get previous page: No client or config available.");
        }

        const prevPage = this.prev_page;
        // TODO: Callbacks should also be propagated to the new instance
        return new (class extends BaseSchemaEndpointConfigurer<Config> implements SchemaEndpointConfigurer<Config> {})(
            this._client,
            {
                ...this._config,
                path: prevPage,
            },
            this._callbacks
        );
    }
}

export async function convertToNavigatablePagedCollection<
    Config extends SchemaEndpointConfig,
    Ref extends HyperlinkEntity,
    Item
>(
    response: unknown,
    client: Client,
    config: Config,
    callbacks: Callbacks<Config>,
    path: string,
    filter?: (ref: Ref) => boolean,
    mapper?: (ref: Ref) => Promise<Item>
): Promise<RTNavigatablePagedCollection<Config, Item>> {
    if (!response || typeof response !== "object") {
        return Promise.reject(new Error(`Invalid response format for resource '${path}'`));
    }
    const page = response as RTPagedCollection<Ref>;
    const refs = page.items ?? [];
    const queues: Item[] = await Promise.all(
        refs
            .filter((ref) => (filter ? filter(ref) : true))
            .map((ref) => (mapper ? mapper(ref) : Promise.resolve(ref as unknown as Item)))
    );
    const newPage: RTPagedCollection<Item> = {
        ...page,
        items: queues,
    };
    return createRTNavigatablePagedCollection<Config, Item>(client, config, callbacks, newPage);
}
