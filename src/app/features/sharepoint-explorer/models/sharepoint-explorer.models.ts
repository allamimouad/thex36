export interface FolderNode {
  isFolder: true;
  name: string;
  serverRelativeUrl: string;
  level?: number;
  children?: FolderNode[];
}

export interface FileItem {
  isFolder: false;
  name: string;
  serverRelativeUrl: string;
  size: number;
  modifiedAt: string;
}

export type ExplorerRow =
  | { kind: 'folder'; folder: FolderNode }
  | { kind: 'file'; file: FileItem };

/** A file extracted from a native drop, with its relative path inside the dropped folder tree. */
export interface UploadEntry {
  /** The file to upload. */
  file: File;
  /**
   * Path relative to the drop target, e.g. "docs/sub/readme.txt".
   * Empty string for files dropped directly (not inside a folder).
   */
  relativePath: string;
}
