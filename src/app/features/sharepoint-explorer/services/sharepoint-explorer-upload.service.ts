import { Injectable } from '@angular/core';
import { delay, Observable, of } from 'rxjs';
import { UploadEntry } from '../models/sharepoint-explorer.models';

@Injectable({ providedIn: 'root' })
export class SharepointExplorerUploadService {
  async extractFiles(event: DragEvent): Promise<UploadEntry[]> {
    const items = event.dataTransfer?.items;
    if (!items) {
      return Array.from(event.dataTransfer?.files ?? []).map((file) => ({ file, relativePath: file.name }));
    }

    const entries: FileSystemEntry[] = [];
    for (let i = 0; i < items.length; i++) {
      const entry = items[i].webkitGetAsEntry?.();
      if (entry) {
        entries.push(entry);
      }
    }

    if (entries.length === 0) {
      return Array.from(event.dataTransfer?.files ?? []).map((file) => ({ file, relativePath: file.name }));
    }

    const result: UploadEntry[] = [];
    await Promise.all(entries.map((entry) => this.collectFiles(entry, '', result)));
    return result;
  }

  private async collectFiles(entry: FileSystemEntry, parentPath: string, result: UploadEntry[]): Promise<void> {
    const relativePath = parentPath ? `${parentPath}/${entry.name}` : entry.name;

    if (entry.isFile) {
      const file = await new Promise<File>((resolve, reject) =>
        (entry as FileSystemFileEntry).file(resolve, reject),
      );
      result.push({ file, relativePath });
    } else if (entry.isDirectory) {
      const dirReader = (entry as FileSystemDirectoryEntry).createReader();
      const childEntries = await this.readAllEntries(dirReader);
      await Promise.all(childEntries.map((child) => this.collectFiles(child, relativePath, result)));
    }
  }

  private readAllEntries(reader: FileSystemDirectoryReader): Promise<FileSystemEntry[]> {
    return new Promise((resolve, reject) => {
      const results: FileSystemEntry[] = [];
      const readBatch = () => {
        reader.readEntries((entries) => {
          if (entries.length === 0) {
            resolve(results);
          } else {
            results.push(...entries);
            readBatch();
          }
        }, reject);
      };
      readBatch();
    });
  }

  uploadFilesTo(folderServerRelativeUrl: string, entries: UploadEntry[]): Observable<void> {
    if (entries.length === 0) {
      return of(void 0);
    }

    // Replace this stub with the real API call.
    // Use entry.relativePath to recreate the folder structure on SharePoint.
    console.log('Upload files', {
      folderServerRelativeUrl,
      files: entries.map((e) => e.relativePath),
    });

    return of(void 0).pipe(delay(1500));
  }

  buildUploadSummary(folderServerRelativeUrl: string, entries: UploadEntry[]): string {
    const folders = new Set(
      entries.map((e) => e.relativePath).filter((p) => p.includes('/')).map((p) => p.split('/')[0]),
    );
    if (folders.size > 0) {
      const folderLabel = folders.size === 1 ? 'folder' : 'folders';
      return `Dropped ${folders.size} ${folderLabel} (${entries.length} files) on ${folderServerRelativeUrl}`;
    }
    const fileLabel = entries.length === 1 ? 'file' : 'files';
    return `Dropped ${entries.length} ${fileLabel} on ${folderServerRelativeUrl}`;
  }
}
