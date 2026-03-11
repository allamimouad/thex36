import { Injectable } from '@angular/core';
import { BehaviorSubject, Observable, map, tap, timer } from 'rxjs';

import { SharepointExplorerClient } from './sharepoint-explorer.client';
import { FileItem, FolderNode } from './sharepoint-explorer.models';

@Injectable({ providedIn: 'root' })
export class SimulatedSharepointExplorerClientService implements SharepointExplorerClient {
  private readonly simulatedLatencyMs = 400;

  private readonly foldersSubject = new BehaviorSubject<FolderNode[]>([
    { isFolder: true, name: 'Root', serverRelativeUrl: '/sites/XXX1/D1' },
    { isFolder: true, name: 'Documents', serverRelativeUrl: '/sites/XXX1/D1/Documents' },
    { isFolder: true, name: 'Images', serverRelativeUrl: '/sites/XXX1/D1/Images' },
    { isFolder: true, name: 'Projects', serverRelativeUrl: '/sites/XXX1/D1/Documents/Projects' },
    { isFolder: true, name: 'Archive', serverRelativeUrl: '/sites/XXX1/D1/Documents/Archive' },
    { isFolder: true, name: 'Mockups', serverRelativeUrl: '/sites/XXX1/D1/Images/Mockups' },
  ]);

  private readonly filesSubject = new BehaviorSubject<FileItem[]>([
    {
      isFolder: false,
      name: 'requirements.docx',
      serverRelativeUrl: '/sites/XXX1/D1/Documents/requirements.docx',
      size: 45682,
      modifiedAt: '2026-03-01',
    },
    {
      isFolder: false,
      name: 'q1-report.xlsx',
      serverRelativeUrl: '/sites/XXX1/D1/Documents/q1-report.xlsx',
      size: 128812,
      modifiedAt: '2026-02-24',
    },
    {
      isFolder: false,
      name: 'hero-banner.png',
      serverRelativeUrl: '/sites/XXX1/D1/Images/hero-banner.png',
      size: 845321,
      modifiedAt: '2026-02-18',
    },
    {
      isFolder: false,
      name: 'mobile-wireframe.fig',
      serverRelativeUrl: '/sites/XXX1/D1/Images/Mockups/mobile-wireframe.fig',
      size: 2251880,
      modifiedAt: '2026-03-04',
    },
    {
      isFolder: false,
      name: 'roadmap.md',
      serverRelativeUrl: '/sites/XXX1/D1/Documents/Projects/roadmap.md',
      size: 8120,
      modifiedAt: '2026-03-05',
    },
  ]);

  getFolderByServerRelativeUrl(folderUrl: string): Observable<FolderNode | null> {
    return this.foldersSubject.pipe(
      map((folders) => folders.find((folder) => folder.serverRelativeUrl === folderUrl) ?? null),
    );
  }

  getFoldersOf(folderUrl: string): Observable<FolderNode[]> {
    return this.foldersSubject.pipe(
      map((folders) =>
        folders.filter((folder) => this.getParentFolderUrl(folder.serverRelativeUrl) === folderUrl),
      ),
    );
  }

  getFilesOf(folderUrl: string): Observable<FileItem[]> {
    return this.filesSubject.pipe(
      map((files) => files.filter((file) => this.getParentFolderUrl(file.serverRelativeUrl) === folderUrl)),
    );
  }

  moveFileTo(fileServerRelativeUrl: string, destinationServerRelativeUrl: string): Observable<void> {
    return timer(this.simulatedLatencyMs).pipe(
      tap(() => {
        const files = this.filesSubject.getValue().map((item) =>
          item.serverRelativeUrl === fileServerRelativeUrl
            ? {
                ...item,
                serverRelativeUrl: destinationServerRelativeUrl,
              }
            : item,
        );

        this.filesSubject.next(files);
      }),
      map(() => void 0),
    );
  }

  moveFolderTo(folderServerRelativeUrl: string, destinationServerRelativeUrl: string): Observable<void> {
    return timer(this.simulatedLatencyMs).pipe(
      tap(() => {
        const folders = this.foldersSubject.getValue().map((folder) => {
          if (
            folder.serverRelativeUrl === folderServerRelativeUrl ||
            folder.serverRelativeUrl.startsWith(`${folderServerRelativeUrl}/`)
          ) {
            return {
              ...folder,
              serverRelativeUrl: folder.serverRelativeUrl.replace(
                folderServerRelativeUrl,
                destinationServerRelativeUrl,
              ),
            };
          }

          return folder;
        });

        const files = this.filesSubject.getValue().map((file) => {
          if (file.serverRelativeUrl.startsWith(`${folderServerRelativeUrl}/`)) {
            return {
              ...file,
              serverRelativeUrl: file.serverRelativeUrl.replace(
                folderServerRelativeUrl,
                destinationServerRelativeUrl,
              ),
            };
          }

          return file;
        });

        this.foldersSubject.next(folders);
        this.filesSubject.next(files);
      }),
      map(() => void 0),
    );
  }

  private getItemName(serverRelativeUrl: string): string {
    return serverRelativeUrl.slice(serverRelativeUrl.lastIndexOf('/') + 1);
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
