import { IFile, IReadableStream } from '@crewdle/web-sdk-types';
import { SharepointReadableStream } from './SharepointReadableStream';

export class SharepointFile implements IFile {
  constructor(public name: string, public size: number, public type: string, public lastModified: number, private file: Blob) {}

  arrayBuffer(): Promise<ArrayBuffer> {
    return this.file.arrayBuffer();
  }

  text(): Promise<string> {
    return this.file.text();
  }

  stream(): IReadableStream {
    return new SharepointReadableStream(this.file.stream());
  }

  slice(start?: number, end?: number, contentType?: string): Blob {
    return this.file.slice(start, end, contentType);
  }
}
