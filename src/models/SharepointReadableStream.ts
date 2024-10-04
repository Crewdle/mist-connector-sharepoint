import { IReadableStream } from '@crewdle/web-sdk-types';

export class SharepointReadableStream implements IReadableStream {
  constructor(private stream: ReadableStream<Uint8Array>) {
    this.stream = stream;
  }

  async read(): Promise<Uint8Array | null> {
    const { value, done } = await this.stream.getReader().read();
    return done ? null : value;
  }

  async close(): Promise<void> {
    this.stream.cancel();
  }
}
