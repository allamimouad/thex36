export interface FolderNode {
  id: string;
  name: string;
  parentId: string | null;
  level?: number;
  children?: FolderNode[];
}

export interface FileItem {
  id: string;
  name: string;
  size: number;
  modifiedAt: string;
  folderId: string;
}

export type ExplorerRow =
  | { kind: 'folder'; folder: FolderNode }
  | { kind: 'file'; file: FileItem };
