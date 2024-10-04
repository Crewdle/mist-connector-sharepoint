import { IFile, IReadableStream } from '@crewdle/web-sdk-types';
export declare class SharepointFile implements IFile {
    name: string;
    size: number;
    type: string;
    lastModified: number;
    private file;
    constructor(name: string, size: number, type: string, lastModified: number, file: Blob);
    arrayBuffer(): Promise<ArrayBuffer>;
    text(): Promise<string>;
    stream(): IReadableStream;
    slice(start?: number, end?: number, contentType?: string): Blob;
}
