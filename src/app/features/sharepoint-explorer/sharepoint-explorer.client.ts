import { Observable } from 'rxjs';

import { FileItem, FolderNode } from './sharepoint-explorer.models';

export abstract class SharepointExplorerClient {
  abstract watchFolders(): Observable<FolderNode[]>;
  abstract watchFiles(): Observable<FileItem[]>;
  abstract moveFileTo(fileServerRelativeUrl: string, destinationFolderUrl: string): Observable<void>;
  abstract moveFolderTo(folderServerRelativeUrl: string, destinationFolderUrl: string): Observable<void>;
}
