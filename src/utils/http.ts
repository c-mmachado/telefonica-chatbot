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

    /** `application/json` content type */
    Json = "application/json",

    /**`text/html` content type */
    Html = "text/html",
}

/**
 * HTTP methods
 *
 * @public
 */
export enum HttpMethods {
    /** HTTP GET method */
    Get = "GET",

    /**  HTTP POST method */
    Post = "POST",

    /** HTTP PUT method */
    Put = "PUT",

    /** HTTP PATCH method */
    Patch = "PATCH",

    /** HTTP DELETE method */
    Delete = "DELETE",
}
