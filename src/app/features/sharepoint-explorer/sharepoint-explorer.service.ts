import { inject, Injectable } from '@angular/core';
import { Observable, combineLatest, map } from 'rxjs';

import { SharepointExplorerClient } from './sharepoint-explorer.client';
import { ExplorerRow, FileItem, FolderNode } from './sharepoint-explorer.models';

@Injectable({ providedIn: 'root' })
export class SharepointExplorerService {
  private readonly client = inject(SharepointExplorerClient);

  watchFolders(): Observable<FolderNode[]> {
    return this.client.watchFolders();
  }

  getRootFolder(): Observable<FolderNode | null> {
    return this.watchFolders().pipe(
      map((folders) => {
        if (folders.length === 0) {
          return null;
        }

        return folders.reduce((root, folder) =>
          folder.serverRelativeUrl.length < root.serverRelativeUrl.length ? folder : root,
        );
      }),
    );
  }

  getFolderByServerRelativeUrl(folderUrl: string): Observable<FolderNode | null> {
    return this.watchFolders().pipe(
      map((folders) => folders.find((folder) => folder.serverRelativeUrl === folderUrl) ?? null),
    );
  }

  getFoldersOf(folderUrl: string): Observable<FolderNode[]> {
    return this.watchFolders().pipe(
      map((folders) =>
        folders.filter((folder) => this.getParentFolderUrl(folder.serverRelativeUrl) === folderUrl),
      ),
    );
  }

  getFolderContentOf(folderUrl: string): Observable<ExplorerRow[]> {
    return combineLatest([this.getFoldersOf(folderUrl), this.client.watchFiles()]).pipe(
      map(([folders, files]) => [
        ...folders.map((folder) => ({ kind: 'folder' as const, folder })),
        ...files
          .filter((file) => this.getParentFolderUrl(file.serverRelativeUrl) === folderUrl)
          .map((file) => ({ kind: 'file' as const, file })),
      ]),
    );
  }

  getFolderPath(folderUrl: string): Observable<FolderNode[]> {
    return this.watchFolders().pipe(
      map((folders) => {
        const path: FolderNode[] = [];
        let current = folders.find((folder) => folder.serverRelativeUrl === folderUrl) ?? null;

        while (current) {
          path.unshift(current);
          const parentFolderUrl = this.getParentFolderUrl(current.serverRelativeUrl);
          current = parentFolderUrl
            ? (folders.find((folder) => folder.serverRelativeUrl === parentFolderUrl) ?? null)
            : null;
        }

        return path;
      }),
    );
  }

  moveFileTo(fileServerRelativeUrl: string, destinationFolderUrl: string): Observable<void> {
    return this.client.moveFileTo(fileServerRelativeUrl, destinationFolderUrl);
  }

  moveFolderTo(folderServerRelativeUrl: string, destinationFolderUrl: string): Observable<void> {
    return this.client.moveFolderTo(folderServerRelativeUrl, destinationFolderUrl);
  }

  private getParentFolderUrl(serverRelativeUrl: string): string | null {
    const lastSlashIndex = serverRelativeUrl.lastIndexOf('/');
    if (lastSlashIndex <= 0) {
      return null;
    }

    const parentFolderUrl = serverRelativeUrl.slice(0, lastSlashIndex);
    return parentFolderUrl.length > 0 ? parentFolderUrl : null;
  }
}
