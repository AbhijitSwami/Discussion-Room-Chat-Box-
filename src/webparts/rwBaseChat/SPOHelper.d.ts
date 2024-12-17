/**
 * Retrieves the Request Digest.
 */
export declare const GetDigest: (url: string) => Promise<any>;
/**
 * Get request for SharePoint API.
 */
export declare const SPGet: (url: string) => Promise<any>;
/**
 * Post request with Request Digest.
 */
export declare const SPPost: ({ url, payload, hdrs, digest, isBlobOrArrayBuffer }: {
    url: string;
    payload?: any;
    hdrs?: Record<string, any> | undefined;
    digest?: any;
    isBlobOrArrayBuffer?: boolean | undefined;
}) => Promise<any>;
/**
 * Uploads a file as an attachment.
 */
export declare const SPFileUpload: ({ url, payload, hdrs, digest }: {
    url: string;
    payload: Blob;
    hdrs?: Record<string, any> | undefined;
    digest?: any;
}) => Promise<any>;
/**
 * Uploads multiple files to SharePoint.
 */
export declare const SPMultiFileUpload: ({ url, hdrs, files, digest }: {
    url: string;
    hdrs?: Record<string, any> | undefined;
    files: Array<{
        fileName: string;
        data: Blob;
    }>;
    digest?: any;
}) => Promise<any[]>;
/**
 * Update request with Request Digest.
 */
export declare const SPUpdate: ({ url, payload, digest }: {
    url: string;
    payload?: any;
    digest?: any;
}) => Promise<any>;
/**
 * Delete request with Request Digest.
 */
export declare const SPDelete: (url: string, digest?: {
    FormDigestValue: string;
}) => Promise<any>;
/**
 * Generates a unique file name with current timestamp.
 */
 sourceMappingURL=SPOHelper.d.ts.map