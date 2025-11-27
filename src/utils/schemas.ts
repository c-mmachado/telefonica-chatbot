import { z } from "zod";

export function parseBoolean(
  value: unknown,
  defaultValue: boolean = false
): boolean {
  if (value === null || value === undefined || value === "") {
    return defaultValue;
  }
  if (typeof value === "boolean") {
    return value;
  }
  if (typeof value === "string") {
    const normalized = value.toLowerCase().trim();
    return (
      normalized === "true" ||
      normalized === "1" ||
      normalized === "yes" ||
      normalized === "on"
    );
  }
  return defaultValue;
}

// Generic array parsing function
export function parseArray<T>(
  value: string | undefined,
  options?: {
    separator?: string;
    defaultValue?: T[];
    parser?: (item: string) => T;
  }
): T[] {
  if (value === null || value === undefined) {
    return options?.defaultValue ?? [];
  }
  if (Array.isArray(value)) {
    return value;
  }
  return value
    .split(options?.separator ?? ",")
    .map((item: string) =>
      options?.parser ? options.parser(item.trim()) : (item.trim() as T)
    );
}

export function parseInteger(
  value: unknown,
  defaultValue: number
): number | unknown {
  if (value === undefined) {
    return defaultValue;
  }
  if (value === "") {
    return NaN;
  }

  let parsed: number | unknown = value;
  if (typeof parsed === "string") {
    parsed = parseInt(parsed, 10);
  }
  if (
    typeof parsed !== "number" ||
    isNaN(parsed) ||
    !Number.isInteger(parsed)
  ) {
    return value; // Let Zod handle the error
  }
  return parsed;
}

// Utility function for boolean environment variables
export const envBoolean = (defaultValue: boolean = false) =>
  z.preprocess((val: unknown): boolean => {
    return parseBoolean(val, defaultValue);
  }, z.boolean());

/**  Utility function for numeric environment variables using Zod. */
export const envInt = (options?: {
  min?: number;
  max?: number;
  default?: number;
}) =>
  z.preprocess(
    (val: unknown): number | unknown => {
      return parseInteger(val, options?.default ?? NaN);
    },
    z
      .number()
      .int()
      .min(options?.min ?? Number.MIN_SAFE_INTEGER)
      .max(options?.max ?? Number.MAX_SAFE_INTEGER)
  );

// Utility function for array environment variables
export function envArray<TZod extends z.ZodType>(
  itemSchema: TZod,
  options?: {
    separator?: string;
    defaultValue?: z.infer<TZod>[];
    parser?: (item: string) => z.infer<TZod>;
  }
) {
  //z.ZodPipe<z.ZodTransform<z.infer<TZod>[], unknown>, z.ZodArray<TZod>>
  return z.preprocess((val: unknown): z.infer<TZod>[] => {
    if (val === null || val === undefined) {
      return options?.defaultValue ?? [];
    }
    if (Array.isArray(val)) {
      return val;
    }
    if (typeof val === "string") {
      return parseArray<z.infer<TZod>>(val, {
        separator: options?.separator,
        defaultValue: options?.defaultValue,
        parser: options?.parser,
      });
    }
    return options?.defaultValue ?? [];
  }, z.array(itemSchema));
}
