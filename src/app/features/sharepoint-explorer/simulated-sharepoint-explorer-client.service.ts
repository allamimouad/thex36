import { Injectable } from '@angular/core';
import { BehaviorSubject, Observable, map, tap, timer } from 'rxjs';

import { SharepointExplorerClient } from './sharepoint-explorer.client';
import { FileItem, FolderNode } from './sharepoint-explorer.models';

@Injectable({ providedIn: 'root' })
export class SimulatedSharepointExplorerClientService implements SharepointExplorerClient {
  private readonly simulatedLatencyMs = 400;

  private readonly foldersSubject = new BehaviorSubject<FolderNode[]>([
    { isFolder: true, name: 'Root', serverRelativeUrl: '/sites/thex36' },
    { isFolder: true, name: 'Documents', serverRelativeUrl: '/sites/thex36/Documents' },
    { isFolder: true, name: 'Images', serverRelativeUrl: '/sites/thex36/Images' },
    { isFolder: true, name: 'Projects', serverRelativeUrl: '/sites/thex36/Documents/Projects' },
    { isFolder: true, name: 'Archive', serverRelativeUrl: '/sites/thex36/Documents/Archive' },
    { isFolder: true, name: 'Mockups', serverRelativeUrl: '/sites/thex36/Images/Mockups' },
  ]);

  private readonly filesSubject = new BehaviorSubject<FileItem[]>([
    {
      isFolder: false,
      name: 'requirements.docx',
      serverRelativeUrl: '/sites/thex36/Documents/requirements.docx',
      size: 45682,
      modifiedAt: '2026-03-01',
    },
    {
      isFolder: false,
      name: 'q1-report.xlsx',
      serverRelativeUrl: '/sites/thex36/Documents/q1-report.xlsx',
      size: 128812,
      modifiedAt: '2026-02-24',
    },
    {
      isFolder: false,
      name: 'hero-banner.png',
      serverRelativeUrl: '/sites/thex36/Images/hero-banner.png',
      size: 845321,
      modifiedAt: '2026-02-18',
    },
    {
      isFolder: false,
      name: 'mobile-wireframe.fig',
      serverRelativeUrl: '/sites/thex36/Images/Mockups/mobile-wireframe.fig',
      size: 2251880,
      modifiedAt: '2026-03-04',
    },
    {
      isFolder: false,
      name: 'roadmap.md',
      serverRelativeUrl: '/sites/thex36/Documents/Projects/roadmap.md',
      size: 8120,
      modifiedAt: '2026-03-05',
    },
  ]);

  watchFolders(): Observable<FolderNode[]> {
    return this.foldersSubject.asObservable();
  }

  watchFiles(): Observable<FileItem[]> {
    return this.filesSubject.asObservable();
  }

  moveFileTo(fileServerRelativeUrl: string, destinationFolderUrl: string): Observable<void> {
    return timer(this.simulatedLatencyMs).pipe(
      tap(() => {
        const files = this.filesSubject.getValue().map((item) =>
          item.serverRelativeUrl === fileServerRelativeUrl
            ? {
                ...item,
                serverRelativeUrl: `${destinationFolderUrl}/${item.name}`,
              }
            : item,
        );

        this.filesSubject.next(files);
      }),
      map(() => void 0),
    );
  }

  moveFolderTo(folderServerRelativeUrl: string, destinationFolderUrl: string): Observable<void> {
    return timer(this.simulatedLatencyMs).pipe(
      tap(() => {
        const nextFolderUrl = this.buildChildUrl(
          destinationFolderUrl,
          this.getItemName(folderServerRelativeUrl),
        );
        const folders = this.foldersSubject.getValue().map((folder) => {
          if (
            folder.serverRelativeUrl === folderServerRelativeUrl ||
            folder.serverRelativeUrl.startsWith(`${folderServerRelativeUrl}/`)
          ) {
            return {
              ...folder,
              serverRelativeUrl: folder.serverRelativeUrl.replace(folderServerRelativeUrl, nextFolderUrl),
            };
          }

          return folder;
        });

        const files = this.filesSubject.getValue().map((file) => {
          if (file.serverRelativeUrl.startsWith(`${folderServerRelativeUrl}/`)) {
            return {
              ...file,
              serverRelativeUrl: file.serverRelativeUrl.replace(folderServerRelativeUrl, nextFolderUrl),
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

  private buildChildUrl(parentFolderUrl: string, childName: string): string {
    return `${parentFolderUrl}/${childName}`;
  }
}
