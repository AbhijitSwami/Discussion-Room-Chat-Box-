/**
 * Checks if an object is of JSON type.
 */
const isObject = (obj: any): boolean => obj !== undefined && obj !== null && obj.constructor === Object;

/**
 * Base HTTP client wrapper around fetch.
 */
const BaseClient = async (url: string, options: RequestInit = {}): Promise<Response> => {
    try {
        return await fetch(url, options);
    } catch (error) {
        console.error("Error in BaseClient:", error);
        throw error;
    }
};

/**
 * Fetches JSON response using BaseClient.
 */
const GetJson = async (url: string, options: RequestInit = {}): Promise<any> => {
    const response = await BaseClient(url, options);
    return response.json();
};

const JSFY = JSON.stringify;

const _headers = {
    credentials: "include",
    Accept: "application/json; odata=nometadata",
    "Content-Type": "application/json; odata=nometadata"
};

let defaultDigest = { FormDigestValue: "" };

const copyObj = (obj: Record<string, any>): Record<string, any> => ({ ...obj });

const MergeObj = (obj1: Record<string, any>, obj2: Record<string, any>): Record<string, any> => ({ ...obj1, ...obj2 });

const _Post = async ({ url, payload = {}, hdrs = {}, isBlobOrArrayBuffer = false }: 
    { url: string, payload?: any, hdrs?: Record<string, any>, isBlobOrArrayBuffer?: boolean }): Promise<any> => {
    const headers = MergeObj(_headers, hdrs);
    const body = isBlobOrArrayBuffer ? payload : JSFY(payload);
    const response = await BaseClient(url, { method: "POST", body, headers });
    if (headers["IF-MATCH"]) return response;
    return response.json();
};

/**
 * Retrieves the Request Digest.
 */
export const GetDigest = async (url: string): Promise<any> => {
    if (!url || typeof url !== "string") throw new Error("Invalid URL");

    const urlSegments = url.toLowerCase().split("/_api");
    if (urlSegments.length > 1) url = `${urlSegments[0]}/_api/contextinfo`;

    return _Post({ url });
};

/**
 * Get request for SharePoint API.
 */
export const SPGet = async (url: string): Promise<any> => {
    const headers = copyObj(_headers);
    return GetJson(url, { headers });
};

/**
 * Post request with Request Digest.
 */
export const SPPost = async ({ url, payload = {}, hdrs = {}, digest = defaultDigest, isBlobOrArrayBuffer = false }: 
    { url: string, payload?: any, hdrs?: Record<string, any>, digest?: any, isBlobOrArrayBuffer?: boolean }): Promise<any> => {
    if (!digest || (isObject(digest) && !digest.FormDigestValue)) digest = await GetDigest(url);
    hdrs["X-RequestDigest"] = digest.FormDigestValue;
    return _Post({ url, hdrs, payload, isBlobOrArrayBuffer });
};

/**
 * Uploads a file as an attachment.
 */
export const SPFileUpload = async ({ url, payload, hdrs = {}, digest = defaultDigest }: 
    { url: string, payload: Blob, hdrs?: Record<string, any>, digest?: any }): Promise<any> => {
    return SPPost({ isBlobOrArrayBuffer: true, url, hdrs, digest, payload });
};

/**
 * Uploads multiple files to SharePoint.
 */
export const SPMultiFileUpload = async ({ url, hdrs = {}, files = [], digest = defaultDigest }: 
    { url: string, hdrs?: Record<string, any>, files: Array<{ fileName: string, data: Blob }>, digest?: any }): Promise<any[]> => {
    const promises = files.map(file =>
        SPFileUpload({ url: `${url}add(FileName='${file.fileName}',overwrite='true')`, payload: file.data, hdrs, digest })
    );
    return Promise.all(promises);
};

/**
 * Update request with Request Digest.
 */
export const SPUpdate = async ({ url, payload = {}, digest = defaultDigest }: 
    { url: string, payload?: any, digest?: any }): Promise<any> => {
    const hdrs = { "IF-MATCH": "*", "X-HTTP-Method": "MERGE" };
    return SPPost({ url, hdrs, payload, digest });
};

/**
 * Delete request with Request Digest.
 */
export const SPDelete = async (url: string, digest = defaultDigest): Promise<any> => {
    const hdrs = { "IF-MATCH": "*", "X-HTTP-Method": "DELETE" };
    return SPPost({ url, hdrs, digest });
};

/**
 * Generates a unique file name with current timestamp.
 */
// const GetUniqueFileName = (): string => `${new Date().getTime()}.txt`;
