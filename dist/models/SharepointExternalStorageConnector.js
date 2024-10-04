import { FileStatus, ObjectKind, StorageEventType } from '@crewdle/web-sdk-types';
import { SharepointFile } from './SharepointFile';
export class SharepointExternalStorageConnector {
    siteId;
    driveId;
    accessToken;
    refreshToken;
    deltaLink;
    constructor(siteId, driveId, accessToken, refreshToken) {
        this.siteId = siteId;
        this.driveId = driveId;
        this.accessToken = accessToken;
        this.refreshToken = refreshToken;
    }
    async get(path, retry = true) {
        try {
            const itemPath = path.split('/').filter((v) => v.length > 0).map(encodeURIComponent).join('/');
            const response = await fetch(`https://graph.microsoft.com/v1.0/sites/${this.siteId}/drives/${this.driveId}/root:/${itemPath}`, {
                headers: {
                    Authorization: `Bearer ${this.accessToken}`,
                },
            });
            if (response.status === 401 && retry) {
                await this.handleRefreshToken();
                return this.get(path, false);
            }
            if (!response.ok) {
                throw new Error(`Failed to get file: ${response.statusText}`);
            }
            const item = await response.json();
            const fileResponse = await fetch(item['@microsoft.graph.downloadUrl']);
            if (!fileResponse.ok) {
                throw new Error(`Failed to get file: ${fileResponse.statusText}`);
            }
            const file = await fileResponse.blob();
            return new SharepointFile(item.name, item.size, item.file.mimeType, item.lastModifiedDateTime, file);
        }
        catch (error) {
            throw new Error(`Failed to get file: ${error.message}`);
        }
    }
    async list(path, recursive, retry = true) {
        const items = [];
        try {
            const itemPath = path.split('/').filter((v) => v.length > 0).map(encodeURIComponent).join('/');
            let url = `https://graph.microsoft.com/v1.0/sites/${this.siteId}/drives/${this.driveId}/root:/${itemPath}:/children`;
            if (itemPath === '') {
                url = `https://graph.microsoft.com/v1.0/sites/${this.siteId}/drives/${this.driveId}/root/children`;
            }
            const response = await fetch(url, {
                headers: {
                    Authorization: `Bearer ${this.accessToken}`,
                },
            });
            if (response.status === 401 && retry) {
                await this.handleRefreshToken();
                return this.list(path, recursive, false);
            }
            if (!response.ok) {
                throw new Error(`Failed to list files: ${response.statusText}`);
            }
            const data = await response.json();
            for (const item of data.value) {
                if (item.folder) {
                    items.push({
                        kind: ObjectKind.Folder,
                        name: item.name,
                        path,
                        pathName: `${path}/${item.name}`,
                        entries: recursive ? await this.list(`${path}/${item.name}`, recursive) : [],
                    });
                }
                else {
                    items.push({
                        kind: ObjectKind.File,
                        name: item.name,
                        path,
                        pathName: `${path}/${item.name}`,
                        size: item.size,
                        status: FileStatus.Synced,
                        type: item.file.mimeType,
                    });
                }
            }
            return items;
        }
        catch (error) {
            throw new Error(`Failed to list files: ${error.message}`);
        }
    }
    async listChanges() {
        const events = [];
        try {
            const url = this.deltaLink ? this.deltaLink : `https://graph.microsoft.com/v1.0/sites/${this.siteId}/drives/${this.driveId}/root/delta(token='latest')`;
            const response = await fetch(url, {
                headers: {
                    Authorization: `Bearer ${this.accessToken}`,
                },
            });
            if (!response.ok) {
                throw new Error(`Failed to list changes: ${response.statusText}`);
            }
            let data = await response.json();
            events.push(...this.processChanges(data));
            while (data['@odata.nextLink']) {
                const nextResponse = await fetch(data['@odata.nextLink'], {
                    headers: {
                        Authorization: `Bearer ${this.accessToken}`,
                    },
                });
                if (!nextResponse.ok) {
                    throw new Error(`Failed to list changes: ${nextResponse.statusText}`);
                }
                data = await nextResponse.json();
                events.push(...this.processChanges(data));
            }
            this.deltaLink = data['@odata.deltaLink'];
        }
        catch (error) {
            throw new Error(`Failed to list changes: ${error.message}`);
        }
        return events;
    }
    async handleRefreshToken() {
        this.accessToken = await this.refreshToken();
    }
    processChanges(data) {
        const events = [];
        for (const item of data.value) {
            if (item.file) {
                if (item.deleted) {
                    events.push({
                        event: StorageEventType.FileDelete,
                        payload: {
                            name: item.name,
                            path: item.parentReference.path,
                            pathName: item.parentReference.path + '/' + item.name,
                        },
                    });
                }
                else {
                    events.push({
                        event: StorageEventType.FileWrite,
                        payload: {
                            file: {
                                kind: ObjectKind.File,
                                path: item.parentReference.path,
                                type: item.file.mimeType,
                                size: item.size,
                                name: item.name,
                                pathName: item.parentReference.path + '/' + item.name,
                                status: FileStatus.Synced,
                            },
                        },
                    });
                }
            }
        }
        return events;
    }
}
