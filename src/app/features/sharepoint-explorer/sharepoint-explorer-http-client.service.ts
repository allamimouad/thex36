import { inject, Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable, map } from 'rxjs';

import { SharepointExplorerClient } from './sharepoint-explorer.client';
import { FileItem, FolderNode } from './sharepoint-explorer.models';

@Injectable({ providedIn: 'root' })
export class SharepointExplorerHttpClientService implements SharepointExplorerClient {
  private readonly http = inject(HttpClient);

  // CHANGE THIS: point to your real backend base path.
  private readonly baseUrl = '/api/sharepoint-explorer';

  getRootFolders(): Observable<FolderNode[]> {
    // CHANGE THIS: endpoint for the initial root folder collection.
    return this.http
      .get<FolderDto[]>(`${this.baseUrl}/folders/root`)
      .pipe(map((folders) => folders.map((folder) => this.mapFolder(folder))));
  }

  getFolderByServerRelativeUrl(folderUrl: string): Observable<FolderNode | null> {
    // CHANGE THIS: adapt query param / route param style to your API.
    return this.http
      .get<FolderDto | null>(`${this.baseUrl}/folders/by-url`, {
        params: { folderUrl },
      })
      .pipe(map((folder) => (folder ? this.mapFolder(folder) : null)));
  }

  getFoldersOf(folderUrl: string): Observable<FolderNode[]> {
    // CHANGE THIS: endpoint for child folders of a given folder.
    return this.http
      .get<FolderDto[]>(`${this.baseUrl}/folders/children`, {
        params: { folderUrl },
      })
      .pipe(map((folders) => folders.map((folder) => this.mapFolder(folder))));
  }

  getFilesOf(folderUrl: string): Observable<FileItem[]> {
    // CHANGE THIS: endpoint for files inside a given folder.
    return this.http
      .get<FileDto[]>(`${this.baseUrl}/files`, {
        params: { folderUrl },
      })
      .pipe(map((files) => files.map((file) => this.mapFile(file))));
  }

  getFolderPath(folderUrl: string): Observable<FolderNode[]> {
    // CHANGE THIS: preferred if your backend can return the breadcrumb directly.
    return this.http
      .get<FolderDto[]>(`${this.baseUrl}/folders/path`, {
        params: { folderUrl },
      })
      .pipe(map((folders) => folders.map((folder) => this.mapFolder(folder))));
  }

  moveFileTo(fileServerRelativeUrl: string, destinationFolderUrl: string): Observable<void> {
    // CHANGE THIS: adapt payload names if your backend expects different fields.
    return this.http.post<void>(`${this.baseUrl}/files/move`, {
      fileServerRelativeUrl,
      destinationFolderUrl,
    });
  }

  moveFolderTo(folderServerRelativeUrl: string, destinationFolderUrl: string): Observable<void> {
    // CHANGE THIS: adapt payload names if your backend expects different fields.
    return this.http.post<void>(`${this.baseUrl}/folders/move`, {
      folderServerRelativeUrl,
      destinationFolderUrl,
    });
  }

  private mapFolder(folder: FolderDto): FolderNode {
    return {
      isFolder: true,
      name: folder.name,
      serverRelativeUrl: folder.serverRelativeUrl,
    };
  }

  private mapFile(file: FileDto): FileItem {
    return {
      isFolder: false,
      name: file.name,
      serverRelativeUrl: file.serverRelativeUrl,
      size: file.size,
      modifiedAt: file.modifiedAt,
    };
  }
}

interface FolderDto {
  name: string;
  serverRelativeUrl: string;
}

interface FileDto {
  name: string;
  serverRelativeUrl: string;
  size: number;
  modifiedAt: string;
}
