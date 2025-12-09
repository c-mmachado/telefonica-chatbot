/**
 * Utility type to make specific properties of a type required.
 *
 * @template T The original type.
 * @template U The keys of the properties to make required.
 * @returns A new type with the specified properties marked as required.
 *
 *  @example
 * ```ts
 * interface User {
 *   id?: string;
 *   name?: string;
 *   email?: string;
 * }
 *
 * // Makes the 'name' property required
 * type UserWithRequiredName = Required<User, 'name'>;
 */
export type Required<T, U extends keyof T> = T & { [key in U]-?: T[key] };

/**
 * Masks sensitive information in an object based on a key regular expression.
 *
 * This function takes an input object and recursively traverses its properties and if a key matches the provided regular expression,
 * its value is replaced with the specified mask string. This is useful for logging or displaying objects without exposing sensitive
 * information such as passwords or secrets.
 *
 * @param source The source object to be masked.
 * @param keyRegEx A regular expression to match keys that should be masked.
 * @param mask The string to replace matched values with.
 * @returns A **new** object with sensitive information masked.
 */
export const maskObject = (
    source: any,
    keyRegEx: string | RegExp = /(?:password|secret)/gi,
    mask: string = "******"
): any => {
    if (!source) {
        return source;
    }

    const target: any = {};
    if (typeof source !== "object" || Array.isArray(source)) {
        // Return non-object values as is
        return source;
    }

    for (const key of Object.keys(source)) {
        if (key.match(keyRegEx)) {
            // Mask matched keys
            target[key] = mask;
            continue;
        }

        if (typeof source[key] === "object") {
            // Recursively mask nested objects
            target[key] = maskObject(source[key], keyRegEx);
        } else {
            target[key] = source[key];
        }
    }

    return target;
};

/**
 * Type guard to check if a key exists in an object.
 *
 * @param key The key to check.
 * @param obj The object to check against.
 * @returns True if the key exists in the object, false otherwise.
 */
export function isKeyOf<K extends string, T extends object>(key: K, obj: T): key is K & keyof T {
    return key in obj;
}
