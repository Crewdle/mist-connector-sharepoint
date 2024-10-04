import { FileStatus, ObjectKind } from '@crewdle/web-sdk-types';
import { SharepointFile } from './SharepointFile';
export class SharepointExternalStorageConnector {
    siteId;
    driveId;
    accessToken;
    refreshToken;
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
            const fileResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${this.siteId}/drives/${this.driveId}/items/${item.id}/content`, {
                headers: {
                    Authorization: `Bearer ${this.accessToken}`,
                },
            });
            if (!fileResponse.ok) {
                throw new Error(`Failed to get file: ${fileResponse.statusText}`);
            }
            const file = await fileResponse.blob();
            return new SharepointFile(item.name, item.size, item.file.type, item.lastModifiedDateTime, file);
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
    async handleRefreshToken() {
        this.accessToken = await this.refreshToken();
    }
}
