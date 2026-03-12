import { Injectable } from '@angular/core';
import { Observable, of } from 'rxjs';

@Injectable({ providedIn: 'root' })
export class SharepointExplorerUploadService {
  extractFiles(event: DragEvent): File[] {
    return Array.from(event.dataTransfer?.files ?? []);
  }

  uploadFilesTo(folderServerRelativeUrl: string, files: File[]): Observable<void> {
    if (files.length === 0) {
      return of(void 0);
    }

    // Replace this stub with the real API call.
    console.log('Upload files', {
      folderServerRelativeUrl,
      fileNames: files.map((file) => file.name),
    });

    return of(void 0);
  }

  buildUploadSummary(folderServerRelativeUrl: string, files: File[]): string {
    const fileLabel = files.length === 1 ? 'file' : 'files';
    return `Dropped ${files.length} ${fileLabel} on ${folderServerRelativeUrl}`;
  }
}
