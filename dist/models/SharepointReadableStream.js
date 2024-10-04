export class SharepointReadableStream {
    stream;
    constructor(stream) {
        this.stream = stream;
        this.stream = stream;
    }
    async read() {
        const { value, done } = await this.stream.getReader().read();
        return done ? null : value;
    }
    async close() {
        this.stream.cancel();
    }
}
