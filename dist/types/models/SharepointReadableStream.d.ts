import { IReadableStream } from '@crewdle/web-sdk-types';
export declare class SharepointReadableStream implements IReadableStream {
    private stream;
    constructor(stream: ReadableStream<Uint8Array>);
    read(): Promise<Uint8Array | null>;
    close(): Promise<void>;
}
