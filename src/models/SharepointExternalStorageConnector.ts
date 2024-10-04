import { FileStatus, IExternalStorageConnectionConnector, IFile, ObjectDescriptor, ObjectKind } from '@crewdle/web-sdk-types';
import { SharepointFile } from './SharepointFile';

export class SharepointExternalStorageConnector implements IExternalStorageConnectionConnector {
  constructor(private siteId: string, private driveId: string, private accessToken: string, private refreshToken: () => Promise<string>) {}

  async get(path: string, retry = true): Promise<IFile> {
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
    } catch (error: any) {
      throw new Error(`Failed to get file: ${error.message}`);
    }
  }

  async list(path: string, recursive: boolean, retry = true): Promise<ObjectDescriptor[]> {
    const items: ObjectDescriptor[] = [];
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
        } else {
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
    } catch (error: any) {
      throw new Error(`Failed to list files: ${error.message}`);
    }
  }

  private async handleRefreshToken() {
    this.accessToken = await this.refreshToken();
  }
}
