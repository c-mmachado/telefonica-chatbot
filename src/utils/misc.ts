export type Required<T, U extends keyof T> = T & { [key in U]-?: T[key] };

