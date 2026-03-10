import { ChangeDetectionStrategy, Component, computed, effect, inject, signal } from '@angular/core';
import { toObservable, toSignal } from '@angular/core/rxjs-interop';
import {
  CdkDrag,
  CdkDragDrop,
  CdkDragEnter,
  CdkDragPreview,
  CdkDropList,
  DragDropModule,
} from '@angular/cdk/drag-drop';
import { TableModule } from 'primeng/table';
import { TreeModule } from 'primeng/tree';
import { TreeNode } from 'primeng/api';
import { of, switchMap } from 'rxjs';
import { ExplorerRow, FileItem, FolderNode } from './sharepoint-explorer.models';
import { SharepointExplorerService } from './sharepoint-explorer.service';

@Component({
  selector: 'app-sharepoint-explorer',
  standalone: true,
  imports: [TreeModule, TableModule, DragDropModule, CdkDropList, CdkDrag, CdkDragPreview],
  changeDetection: ChangeDetectionStrategy.OnPush,
  templateUrl: './sharepoint-explorer.component.html',
  styleUrl: './sharepoint-explorer.component.scss',
})
export class SharepointExplorerComponent {
  private readonly explorerService = inject(SharepointExplorerService);

  readonly allFolders = toSignal(this.explorerService.watchFolders(), { initialValue: [] as FolderNode[] });
  readonly rootFolder = toSignal(this.explorerService.getRootFolder(), { initialValue: null });
  readonly selectedFolderId = signal<string>('');
  readonly activeDropFolderId = signal<string | null>(null);
  readonly expandedFolderIds = signal<Set<string>>(new Set<string>());

  readonly treeNodes = computed<TreeNode<FolderNode>[]>(() => {
    const rootFolder = this.rootFolder();
    if (!rootFolder) {
      return [];
    }

    return [this.toTreeNode(rootFolder, 0, this.expandedFolderIds(), this.allFolders())];
  });
  readonly selectedItems = toSignal(
    toObservable(this.selectedFolderId).pipe(
      switchMap((folderId) => (folderId ? this.explorerService.getFolderContentOf(folderId) : of([] as ExplorerRow[]))),
    ),
    { initialValue: [] as ExplorerRow[] },
  );
  readonly breadcrumb = toSignal(
    toObservable(this.selectedFolderId).pipe(
      switchMap((folderId) => (folderId ? this.explorerService.getFolderPath(folderId) : of([] as FolderNode[]))),
    ),
    { initialValue: [] as FolderNode[] },
  );
  readonly selectedFolder = computed(
    () => this.allFolders().find((folder) => folder.id === this.selectedFolderId()) ?? null,
  );

  constructor() {
    effect(() => {
      const rootFolder = this.rootFolder();
      if (!rootFolder || this.selectedFolderId()) {
        return;
      }

      this.selectedFolderId.set(rootFolder.id);
      this.expandedFolderIds.set(new Set([rootFolder.id]));
    });
  }

  onFolderSelect(folderId: string): void {
    this.selectFolder(folderId);
  }

  onOpenFolder(folderId: string): void {
    this.selectFolder(folderId);
  }

  onDropListEntered(event: CdkDragEnter<FolderNode | null, FolderNode | null>): void {
    this.activeDropFolderId.set(event.container.data?.id ?? null);
  }

  onDropListExited(): void {
    this.activeDropFolderId.set(null);
  }

  onTreeNodeExpand(event: { node: TreeNode<FolderNode> }): void {
    const folder = event.node.data;
    if (!folder) {
      return;
    }

    this.expandedFolderIds.update((current) => new Set(current).add(folder.id));
  }

  onTreeNodeCollapse(event: { node: TreeNode<FolderNode> }): void {
    const folder = event.node.data;
    if (!folder) {
      return;
    }

    this.expandedFolderIds.update((current) => {
      const next = new Set(current);
      next.delete(folder.id);
      return next;
    });
  }

  onDropOnFolder(
    event: CdkDragDrop<FolderNode | null, unknown, FileItem | FolderNode>,
    targetFolder: FolderNode,
  ): void {
    this.activeDropFolderId.set(null);
    const dragged = event.item.data;

    if (this.isFileItem(dragged)) {
      this.moveFile(dragged, targetFolder);
      return;
    }

    if (this.isFolderNode(dragged)) {
      this.moveFolder(dragged, targetFolder);
    }
  }

  canEnterFolder = (drag: CdkDrag<FileItem | FolderNode>, drop: CdkDropList<FolderNode>): boolean => {
    return this.canDropIntoFolder(drag.data, drop.data);
  };

  canEnterCurrentFolder = (
    drag: CdkDrag<FileItem | FolderNode>,
    _drop: CdkDropList<FolderNode | null>,
  ): boolean => {
    return this.canDropIntoFolder(drag.data, this.selectedFolder());
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
    let current = this.allFolders().find((folder) => folder.id === candidateId) ?? null;

    while (current?.parentId) {
      if (current.parentId === ancestorId) {
        return true;
      }
      current = this.allFolders().find((folder) => folder.id === current?.parentId) ?? null;
    }

    return false;
  }

  private canDropIntoFolder(dragged: FileItem | FolderNode, targetFolder: FolderNode | null): boolean {
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
  }

  private moveFile(file: FileItem, targetFolder: FolderNode): void {
    if (file.folderId === targetFolder.id) {
      return;
    }

    this.explorerService.moveFileTo(file.id, targetFolder.id).subscribe();
  }

  private moveFolder(folder: FolderNode, targetFolder: FolderNode): void {
    if (!this.isValidFolderMove(folder, targetFolder)) {
      return;
    }

    this.explorerService.moveFolderTo(folder.id, targetFolder.id).subscribe();
  }

  private selectFolder(folderId: string): void {
    this.selectedFolderId.set(folderId);
  }

  private toTreeNode(
    folder: FolderNode,
    level: number,
    expandedFolderIds: Set<string>,
    allFolders: FolderNode[],
  ): TreeNode<FolderNode> {
    const childFolders = allFolders.filter((childFolder) => childFolder.parentId === folder.id);
    const isExpanded = expandedFolderIds.has(folder.id);

    return {
      key: folder.id,
      label: folder.name,
      data: { ...folder, level },
      expanded: isExpanded,
      leaf: childFolders.length === 0,
      children: isExpanded
        ? childFolders.map((childFolder) => this.toTreeNode(childFolder, level + 1, expandedFolderIds, allFolders))
        : [],
    };
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
