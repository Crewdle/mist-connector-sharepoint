import { IExternalStorageConnectionConnector, IFile, ObjectDescriptor, StorageEvent } from '@crewdle/web-sdk-types';
export declare class SharepointExternalStorageConnector implements IExternalStorageConnectionConnector {
    private siteId;
    private driveId;
    private accessToken;
    private refreshToken;
    private deltaLink?;
    constructor(siteId: string, driveId: string, accessToken: string, refreshToken: () => Promise<string>);
    get(path: string, retry?: boolean): Promise<IFile>;
    list(path: string, recursive: boolean, retry?: boolean): Promise<ObjectDescriptor[]>;
    listChanges(retry?: boolean): Promise<StorageEvent[]>;
    private handleRefreshToken;
    private processChanges;
}
