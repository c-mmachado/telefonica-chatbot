export interface RequestBuilder<T> {
    get(): Promise<T>;
}
