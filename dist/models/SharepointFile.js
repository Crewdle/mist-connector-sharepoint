import { SharepointReadableStream } from './SharepointReadableStream';
export class SharepointFile {
    name;
    size;
    type;
    lastModified;
    file;
    constructor(name, size, type, lastModified, file) {
        this.name = name;
        this.size = size;
        this.type = type;
        this.lastModified = lastModified;
        this.file = file;
    }
    arrayBuffer() {
        return this.file.arrayBuffer();
    }
    text() {
        return this.file.text();
    }
    stream() {
        return new SharepointReadableStream(this.file.stream());
    }
    slice(start, end, contentType) {
        return this.file.slice(start, end, contentType);
    }
}
