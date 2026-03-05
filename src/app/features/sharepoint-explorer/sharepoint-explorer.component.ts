import { ChangeDetectionStrategy, Component, computed, signal } from '@angular/core';
import {
  CdkDrag,
  CdkDragDrop,
  CdkDragEnter,
  CdkDragExit,
  CdkDragPreview,
  CdkDropList,
  DragDropModule,
} from '@angular/cdk/drag-drop';
import { TableModule } from 'primeng/table';
import { TreeModule } from 'primeng/tree';
import { TreeNode } from 'primeng/api';

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

type ExplorerRow =
  | { kind: 'folder'; folder: FolderNode }
  | { kind: 'file'; file: FileItem };

class ExplorerApiService {
  apiMoveFile(fileId: string, targetFolderId: string): void {
    console.log('apiMoveFile', { fileId, targetFolderId });
  }

  apiMoveFolder(folderId: string, targetFolderId: string): void {
    console.log('apiMoveFolder', { folderId, targetFolderId });
  }
}

@Component({
  selector: 'app-sharepoint-explorer',
  standalone: true,
  imports: [TreeModule, TableModule, DragDropModule, CdkDropList, CdkDrag, CdkDragPreview],
  changeDetection: ChangeDetectionStrategy.OnPush,
  templateUrl: './sharepoint-explorer.component.html',
  styleUrl: './sharepoint-explorer.component.scss',
})
export class SharepointExplorerComponent {
  private readonly api = new ExplorerApiService();

  private readonly initialFolders: FolderNode[] = [
    { id: 'root', name: 'Root', parentId: null },
    { id: 'docs', name: 'Documents', parentId: 'root' },
    { id: 'images', name: 'Images', parentId: 'root' },
    { id: 'projects', name: 'Projects', parentId: 'docs' },
    { id: 'archive', name: 'Archive', parentId: 'docs' },
    { id: 'mockups', name: 'Mockups', parentId: 'images' },
  ];

  private readonly initialFiles: FileItem[] = [
    { id: 'f1', name: 'requirements.docx', size: 45682, modifiedAt: '2026-03-01', folderId: 'docs' },
    { id: 'f2', name: 'q1-report.xlsx', size: 128812, modifiedAt: '2026-02-24', folderId: 'docs' },
    { id: 'f3', name: 'hero-banner.png', size: 845321, modifiedAt: '2026-02-18', folderId: 'images' },
    { id: 'f4', name: 'mobile-wireframe.fig', size: 2251880, modifiedAt: '2026-03-04', folderId: 'mockups' },
    { id: 'f5', name: 'roadmap.md', size: 8120, modifiedAt: '2026-03-05', folderId: 'projects' },
  ];

  readonly flatFolders = signal<FolderNode[]>(this.initialFolders);
  readonly files = signal<FileItem[]>(this.initialFiles);
  readonly selectedFolderId = signal<string>(this.initialFolders[0]?.id ?? '');
  readonly activeDropFolderId = signal<string | null>(null);

  readonly folderMap = computed(() => {
    const map = new Map<string, FolderNode>();
    for (const folder of this.flatFolders()) {
      map.set(folder.id, folder);
    }
    return map;
  });

  readonly folderTree = computed(() => this.computeFolderTree(this.flatFolders()));
  readonly treeNodes = computed<TreeNode<FolderNode>[]>(() => this.toPrimeTreeNodes(this.folderTree()));

  readonly selectedFiles = computed(() => {
    const folderId = this.selectedFolderId();
    return this.files().filter((file) => file.folderId === folderId);
  });
  readonly selectedChildFolders = computed(() => {
    const folderId = this.selectedFolderId();
    return this.flatFolders().filter((folder) => folder.parentId === folderId);
  });
  readonly selectedItems = computed<ExplorerRow[]>(() => [
    ...this.selectedChildFolders().map((folder) => ({ kind: 'folder' as const, folder })),
    ...this.selectedFiles().map((file) => ({ kind: 'file' as const, file })),
  ]);

  readonly breadcrumb = computed(() => this.getFolderPath(this.selectedFolderId()));
  readonly selectedFolder = computed(() => this.folderMap().get(this.selectedFolderId()) ?? null);

  onFolderSelect(folderId: string): void {
    this.selectedFolderId.set(folderId);
  }

  onOpenFolder(folderId: string): void {
    this.selectedFolderId.set(folderId);
  }

  onDropListEntered(event: CdkDragEnter<FolderNode | null, FolderNode | null>): void {
    this.activeDropFolderId.set(event.container.data?.id ?? null);
  }

  onDropListExited(_event: CdkDragExit<FolderNode | null, FolderNode | null>): void {
    this.activeDropFolderId.set(null);
  }

  onDropOnFolder(
    event: CdkDragDrop<FolderNode | null, unknown, FileItem | FolderNode>,
    targetFolder: FolderNode,
  ): void {
    this.activeDropFolderId.set(null);
    const dragged = event.item.data;

    if (this.isFileItem(dragged)) {
      if (dragged.folderId === targetFolder.id) {
        return;
      }
      this.files.update((items) =>
        items.map((item) => (item.id === dragged.id ? { ...item, folderId: targetFolder.id } : item)),
      );
      this.api.apiMoveFile(dragged.id, targetFolder.id);
      return;
    }

    if (this.isFolderNode(dragged)) {
      if (!this.isValidFolderMove(dragged, targetFolder)) {
        return;
      }
      this.flatFolders.update((items) =>
        items.map((folder) => (folder.id === dragged.id ? { ...folder, parentId: targetFolder.id } : folder)),
      );
      this.api.apiMoveFolder(dragged.id, targetFolder.id);
    }
  }

  canEnterFolder = (drag: CdkDrag<FileItem | FolderNode>, drop: CdkDropList<FolderNode>): boolean => {
    const targetFolder = drop.data;
    const dragged = drag.data;

    if (!targetFolder) {
      return false;
    }

    if (this.isFileItem(dragged)) {
      return dragged.folderId !== targetFolder.id;
    }

    if (this.isFolderNode(dragged)) {
      return this.isValidFolderMove(dragged, targetFolder);
    }

    return false;
  };

  canEnterCurrentFolder = (
    drag: CdkDrag<FileItem | FolderNode>,
    _drop: CdkDropList<FolderNode | null>,
  ): boolean => {
    const targetFolder = this.selectedFolder();
    if (!targetFolder) {
      return false;
    }

    const dragged = drag.data;
    if (this.isFileItem(dragged)) {
      return dragged.folderId !== targetFolder.id;
    }

    if (this.isFolderNode(dragged)) {
      return this.isValidFolderMove(dragged, targetFolder);
    }

    return false;
  };

  private isValidFolderMove(source: FolderNode, target: FolderNode): boolean {
    if (source.id === target.id) {
      return false;
    }
    if (this.isDescendant(target.id, source.id)) {
      return false;
    }
    return source.parentId !== target.id;
  }

  private isDescendant(candidateId: string, ancestorId: string): boolean {
    const map = this.folderMap();
    let current = map.get(candidateId) ?? null;

    while (current?.parentId) {
      if (current.parentId === ancestorId) {
        return true;
      }
      current = map.get(current.parentId) ?? null;
    }
    return false;
  }

  private computeFolderTree(flatFolders: FolderNode[]): FolderNode[] {
    const map = new Map<string, FolderNode>();
    const roots: FolderNode[] = [];

    for (const folder of flatFolders) {
      map.set(folder.id, { ...folder, children: [] });
    }

    for (const folder of flatFolders) {
      const node = map.get(folder.id);
      if (!node) {
        continue;
      }
      if (folder.parentId && map.has(folder.parentId)) {
        map.get(folder.parentId)?.children?.push(node);
      } else {
        roots.push(node);
      }
    }

    return roots;
  }

  private toPrimeTreeNodes(nodes: FolderNode[], level = 0): TreeNode<FolderNode>[] {
    return nodes.map((folder) => ({
      key: folder.id,
      label: folder.name,
      data: {
        id: folder.id,
        name: folder.name,
        parentId: folder.parentId,
        level,
      },
      expanded: true,
      children: this.toPrimeTreeNodes(folder.children ?? [], level + 1),
    }));
  }

  private getFolderPath(folderId: string): FolderNode[] {
    const map = this.folderMap();
    const path: FolderNode[] = [];
    let current = map.get(folderId) ?? null;

    while (current) {
      path.unshift(current);
      current = current.parentId ? (map.get(current.parentId) ?? null) : null;
    }

    return path;
  }

  private isFileItem(value: unknown): value is FileItem {
    return !!value && typeof value === 'object' && 'folderId' in value && 'size' in value;
  }

  private isFolderNode(value: unknown): value is FolderNode {
    return !!value && typeof value === 'object' && 'parentId' in value && 'name' in value;
  }

  formatSize(size: number): string {
    if (size < 1024) {
      return `${size} B`;
    }
    if (size < 1024 * 1024) {
      return `${(size / 1024).toFixed(1)} KB`;
    }
    return `${(size / (1024 * 1024)).toFixed(1)} MB`;
  }
}
