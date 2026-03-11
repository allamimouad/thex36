import { Observable } from 'rxjs';

import { FileItem, FolderNode } from './sharepoint-explorer.models';

export abstract class SharepointExplorerClient {
  abstract getFolderByServerRelativeUrl(folderUrl: string): Observable<FolderNode | null>;
  abstract getFoldersOf(folderUrl: string): Observable<FolderNode[]>;
  abstract getFilesOf(folderUrl: string): Observable<FileItem[]>;
  abstract moveFileTo(fileServerRelativeUrl: string, destinationServerRelativeUrl: string): Observable<void>;
  abstract moveFolderTo(folderServerRelativeUrl: string, destinationServerRelativeUrl: string): Observable<void>;
}
