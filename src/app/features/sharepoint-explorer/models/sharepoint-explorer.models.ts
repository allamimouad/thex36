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
