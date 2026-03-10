import { Injectable } from '@angular/core';
import { BehaviorSubject, Observable, combineLatest, map, of } from 'rxjs';

import { ExplorerRow, FileItem, FolderNode } from './sharepoint-explorer.models';

@Injectable({ providedIn: 'root' })
export class SharepointExplorerService {
  private readonly foldersSubject = new BehaviorSubject<FolderNode[]>([
    { id: 'root', name: 'Root', parentId: null },
    { id: 'docs', name: 'Documents', parentId: 'root' },
    { id: 'images', name: 'Images', parentId: 'root' },
    { id: 'projects', name: 'Projects', parentId: 'docs' },
    { id: 'archive', name: 'Archive', parentId: 'docs' },
    { id: 'mockups', name: 'Mockups', parentId: 'images' },
  ]);

  private readonly filesSubject = new BehaviorSubject<FileItem[]>([
    { id: 'f1', name: 'requirements.docx', size: 45682, modifiedAt: '2026-03-01', folderId: 'docs' },
    { id: 'f2', name: 'q1-report.xlsx', size: 128812, modifiedAt: '2026-02-24', folderId: 'docs' },
    { id: 'f3', name: 'hero-banner.png', size: 845321, modifiedAt: '2026-02-18', folderId: 'images' },
    { id: 'f4', name: 'mobile-wireframe.fig', size: 2251880, modifiedAt: '2026-03-04', folderId: 'mockups' },
    { id: 'f5', name: 'roadmap.md', size: 8120, modifiedAt: '2026-03-05', folderId: 'projects' },
  ]);

  watchFolders(): Observable<FolderNode[]> {
    return this.foldersSubject.asObservable();
  }

  getRootFolder(): Observable<FolderNode | null> {
    return this.watchFolders().pipe(map((folders) => folders.find((folder) => folder.parentId === null) ?? null));
  }

  getFolderById(folderId: string): Observable<FolderNode | null> {
    return this.watchFolders().pipe(map((folders) => folders.find((folder) => folder.id === folderId) ?? null));
  }

  getFoldersOf(folderId: string): Observable<FolderNode[]> {
    return this.watchFolders().pipe(map((folders) => folders.filter((folder) => folder.parentId === folderId)));
  }

  getFolderContentOf(folderId: string): Observable<ExplorerRow[]> {
    return combineLatest([this.getFoldersOf(folderId), this.filesSubject.asObservable()]).pipe(
      map(([folders, files]) => [
        ...folders.map((folder) => ({ kind: 'folder' as const, folder })),
        ...files
          .filter((file) => file.folderId === folderId)
          .map((file) => ({ kind: 'file' as const, file })),
      ]),
    );
  }

  getFolderPath(folderId: string): Observable<FolderNode[]> {
    return this.watchFolders().pipe(
      map((folders) => {
        const path: FolderNode[] = [];
        let current = folders.find((folder) => folder.id === folderId) ?? null;

        while (current) {
          path.unshift(current);
          const parentId = current.parentId;
          current = parentId ? (folders.find((folder) => folder.id === parentId) ?? null) : null;
        }

        return path;
      }),
    );
  }

  moveFileTo(fileId: string, folderId: string): Observable<void> {
    const files = this.filesSubject
      .getValue()
      .map((item) => (item.id === fileId ? { ...item, folderId } : item));

    this.filesSubject.next(files);
    return of(void 0);
  }

  moveFolderTo(folderId: string, destinationFolderId: string): Observable<void> {
    const folders = this.foldersSubject
      .getValue()
      .map((folder) => (folder.id === folderId ? { ...folder, parentId: destinationFolderId } : folder));

    this.foldersSubject.next(folders);
    return of(void 0);
  }
}
