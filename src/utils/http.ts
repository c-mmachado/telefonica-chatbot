/**
 * HTTP headers
 *
 * @public
 */
export enum HttpHeaders {
    /** Content-Type header */
    ContentType = "Content-Type",

    /** Authorization header */
    Authorization = "Authorization",

    /** Accept header */
    Accept = "Accept",

    /** Cookie header */
    Cookie = "Cookie",
}

/**
 * HTTP content types
 *
 * @public
 */
export enum HttpContentTypes {
    /** `application/x-www-form-urlencoded` content type */
    FormUrlEncoded = "application/x-www-form-urlencoded",

    /** `multipart/form-data` content type */
    MultipartFormData = "multipart/form-data",

    /** `application/json` content type */
    Json = "application/json",

    /**`text/html` content type */
    Html = "text/html",

    /** `application/xml` content type */
    Xml = "application/xml",

    /** `text/plain` content type */
    TextPlain = "text/plain",
}

/**
 * HTTP methods
 *
 * @public
 */
export enum HttpMethod {
    /** HTTP GET method */
    Get = "get",

    /**  HTTP POST method */
    Post = "post",

    /** HTTP PUT method */
    Put = "put",

    /** HTTP PATCH method */
    Patch = "patch",

    /** HTTP DELETE method */
    Delete = "delete",
}
