import { inject, Injectable } from '@angular/core';
import { BehaviorSubject, Observable, combineLatest, forkJoin, map, of, switchMap, take, tap } from 'rxjs';

import { SharepointExplorerClient } from './sharepoint-explorer.client';
import { ExplorerRow, FolderNode } from './sharepoint-explorer.models';

export const SHAREPOINT_EXPLORER_ROOT_URL = '/sites/XXX1/D1';

@Injectable({ providedIn: 'root' })
export class SharepointExplorerService {
  private readonly client = inject(SharepointExplorerClient);

  private readonly foldersSubject = new BehaviorSubject<FolderNode[]>([]);
  private readonly refreshSubject = new BehaviorSubject(0);
  private readonly loadedParents = new Set<string | null>();

  watchFolders(): Observable<FolderNode[]> {
    return this.foldersSubject.asObservable();
  }

  getRootFolders(): Observable<FolderNode[]> {
    return this.getFoldersOf(SHAREPOINT_EXPLORER_ROOT_URL);
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
    return this.refreshSubject.pipe(
      switchMap(() =>
        combineLatest([
          this.client.getFoldersOf(folderUrl).pipe(take(1)),
          this.client.getFilesOf(folderUrl).pipe(take(1)),
        ]),
      ),
      map(([folders, files]) => [
        ...folders.map((folder) => ({ kind: 'folder' as const, folder })),
        ...files.map((file) => ({ kind: 'file' as const, file })),
      ]),
    );
  }

  loadRootFolders(): Observable<FolderNode[]> {
    if (this.loadedParents.has(SHAREPOINT_EXPLORER_ROOT_URL)) {
      return this.getRootFolders().pipe(take(1));
    }

    return this.client.getFoldersOf(SHAREPOINT_EXPLORER_ROOT_URL).pipe(
      take(1),
      tap((folders) => {
        this.loadedParents.add(SHAREPOINT_EXPLORER_ROOT_URL);
        this.replaceChildren(SHAREPOINT_EXPLORER_ROOT_URL, folders);
      }),
    );
  }

  loadFoldersOf(folderUrl: string, forceRefresh = false): Observable<FolderNode[]> {
    if (!forceRefresh && this.loadedParents.has(folderUrl)) {
      return this.getFoldersOf(folderUrl).pipe(take(1));
    }

    return this.client.getFoldersOf(folderUrl).pipe(
      take(1),
      tap((folders) => {
        this.loadedParents.add(folderUrl);
        this.replaceChildren(folderUrl, folders);
      }),
    );
  }

  moveFileTo(fileServerRelativeUrl: string, destinationServerRelativeUrl: string): Observable<void> {
    const sourceFolderUrl = this.getParentFolderUrl(fileServerRelativeUrl);
    const destinationFolderUrl = this.getParentFolderUrl(destinationServerRelativeUrl);

    if (!destinationFolderUrl) {
      return of(void 0);
    }

    return this.client.moveFileTo(fileServerRelativeUrl, destinationServerRelativeUrl).pipe(
      switchMap(() => this.reloadAfterMove(sourceFolderUrl, destinationFolderUrl)),
    );
  }

  moveFolderTo(folderServerRelativeUrl: string, destinationServerRelativeUrl: string): Observable<void> {
    const sourceFolderUrl = this.getParentFolderUrl(folderServerRelativeUrl);
    const destinationFolderUrl = this.getParentFolderUrl(destinationServerRelativeUrl);

    if (!destinationFolderUrl) {
      return of(void 0);
    }

    return this.client.moveFolderTo(folderServerRelativeUrl, destinationServerRelativeUrl).pipe(
      switchMap(() => this.reloadAfterMove(sourceFolderUrl, destinationFolderUrl)),
    );
  }

  isFolderLoaded(folderUrl: string | null): boolean {
    return this.loadedParents.has(folderUrl);
  }

  private reloadAfterMove(sourceFolderUrl: string | null, destinationFolderUrl: string): Observable<void> {
    const reloads: Observable<unknown>[] = [];

    if (sourceFolderUrl === SHAREPOINT_EXPLORER_ROOT_URL) {
      reloads.push(this.loadRootFoldersForce());
    } else if (sourceFolderUrl && this.loadedParents.has(sourceFolderUrl)) {
      reloads.push(this.loadFoldersOf(sourceFolderUrl, true));
    }

    if (destinationFolderUrl === SHAREPOINT_EXPLORER_ROOT_URL) {
      reloads.push(this.loadRootFoldersForce());
    } else if (this.loadedParents.has(destinationFolderUrl)) {
      reloads.push(this.loadFoldersOf(destinationFolderUrl, true));
    }

    if (reloads.length === 0) {
      this.refreshSubject.next(this.refreshSubject.value + 1);
      return of(void 0);
    }

    return forkJoin(reloads).pipe(
      tap(() => this.refreshSubject.next(this.refreshSubject.value + 1)),
      map(() => void 0),
    );
  }

  private loadRootFoldersForce(): Observable<FolderNode[]> {
    this.loadedParents.delete(SHAREPOINT_EXPLORER_ROOT_URL);
    return this.loadRootFolders();
  }

  private replaceChildren(parentFolderUrl: string | null, nextChildren: FolderNode[]): void {
    const currentFolders = this.foldersSubject.getValue();
    const currentChildren = currentFolders.filter(
      (folder) => this.getParentFolderUrl(folder.serverRelativeUrl) === parentFolderUrl,
    );

    const removedChildUrls = currentChildren
      .filter(
        (currentChild) =>
          !nextChildren.some((nextChild) => nextChild.serverRelativeUrl === currentChild.serverRelativeUrl),
      )
      .map((folder) => folder.serverRelativeUrl);

    const remainingFolders = currentFolders.filter(
      (folder) =>
        this.getParentFolderUrl(folder.serverRelativeUrl) !== parentFolderUrl &&
        !removedChildUrls.some((removedUrl) => this.isSameOrDescendant(folder.serverRelativeUrl, removedUrl)),
    );

    removedChildUrls.forEach((folderUrl) => this.clearLoadedBranch(folderUrl));

    const nextFolders = [...remainingFolders];
    nextChildren.forEach((folder) => {
      const existingIndex = nextFolders.findIndex(
        (existingFolder) => existingFolder.serverRelativeUrl === folder.serverRelativeUrl,
      );

      if (existingIndex >= 0) {
        nextFolders[existingIndex] = folder;
        return;
      }

      nextFolders.push(folder);
    });

    this.foldersSubject.next(nextFolders);
  }

  private clearLoadedBranch(folderUrl: string): void {
    const branchUrls = this.foldersSubject
      .getValue()
      .filter((folder) => this.isSameOrDescendant(folder.serverRelativeUrl, folderUrl))
      .map((folder) => folder.serverRelativeUrl);

    branchUrls.forEach((url) => {
      this.loadedParents.delete(url);
    });
  }

  private isSameOrDescendant(candidateUrl: string, ancestorUrl: string): boolean {
    return candidateUrl === ancestorUrl || candidateUrl.startsWith(`${ancestorUrl}/`);
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
