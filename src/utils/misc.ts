export type Required<T, U extends keyof T> = T & { [key in U]-?: T[key] };

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

export function isKeyOf<K extends string, T extends object>(key: K, obj: T): key is K & keyof T {
    return key in obj;
}
