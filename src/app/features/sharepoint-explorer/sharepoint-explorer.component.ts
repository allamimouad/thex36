import { ChangeDetectionStrategy, Component, DestroyRef, computed, effect, inject, signal } from '@angular/core';
import { takeUntilDestroyed, toObservable, toSignal } from '@angular/core/rxjs-interop';
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
import { finalize, of, switchMap } from 'rxjs';
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
  private readonly destroyRef = inject(DestroyRef);
  private readonly explorerService = inject(SharepointExplorerService);

  readonly allFolders = toSignal(this.explorerService.watchFolders(), { initialValue: [] as FolderNode[] });
  readonly rootFolders = toSignal(this.explorerService.getRootFolders(), { initialValue: [] as FolderNode[] });
  readonly selectedFolderUrl = signal<string>('');
  readonly activeDropFolderUrl = signal<string | null>(null);
  readonly expandedFolderUrls = signal<Set<string>>(new Set<string>());
  readonly pendingOperations = signal(0);
  readonly isOperating = computed(() => this.pendingOperations() > 0);

  readonly treeNodes = computed<TreeNode<FolderNode>[]>(() => {
    return this.rootFolders().map((rootFolder) =>
      this.toTreeNode(rootFolder, 0, this.expandedFolderUrls(), this.allFolders()),
    );
  });
  readonly selectedItems = toSignal(
    toObservable(this.selectedFolderUrl).pipe(
      switchMap((folderUrl) =>
        folderUrl ? this.explorerService.getFolderContentOf(folderUrl) : of([] as ExplorerRow[])
      ),
    ),
    { initialValue: [] as ExplorerRow[] },
  );
  readonly breadcrumb = toSignal(
    toObservable(this.selectedFolderUrl).pipe(
      switchMap((folderUrl) => (folderUrl ? this.explorerService.getFolderPath(folderUrl) : of([] as FolderNode[]))),
    ),
    { initialValue: [] as FolderNode[] },
  );
  readonly selectedFolder = computed(
    () => this.allFolders().find((folder) => folder.serverRelativeUrl === this.selectedFolderUrl()) ?? null,
  );

  constructor() {
    this.explorerService
      .loadRootFolders()
      .pipe(takeUntilDestroyed(this.destroyRef))
      .subscribe();

    effect(() => {
      const rootFolders = this.rootFolders();
      if (rootFolders.length === 0 || this.selectedFolderUrl()) {
        return;
      }

      this.selectedFolderUrl.set(rootFolders[0].serverRelativeUrl);
    });
  }

  onFolderSelect(folderUrl: string): void {
    this.selectFolder(folderUrl);
  }

  onOpenFolder(folderUrl: string): void {
    this.selectFolder(folderUrl);
  }

  onDropListEntered(event: CdkDragEnter<FolderNode | null, FolderNode | null>): void {
    this.activeDropFolderUrl.set(event.container.data?.serverRelativeUrl ?? null);
  }

  onDropListExited(): void {
    this.activeDropFolderUrl.set(null);
  }

  onTreeNodeExpand(event: { node: TreeNode<FolderNode> }): void {
    const folder = event.node.data;
    if (!folder) {
      return;
    }

    this.expandedFolderUrls.update((current) => new Set(current).add(folder.serverRelativeUrl));
    this.explorerService
      .loadFoldersOf(folder.serverRelativeUrl)
      .pipe(takeUntilDestroyed(this.destroyRef))
      .subscribe();
  }

  onTreeNodeCollapse(event: { node: TreeNode<FolderNode> }): void {
    const folder = event.node.data;
    if (!folder) {
      return;
    }

    this.expandedFolderUrls.update((current) => {
      const next = new Set(current);
      next.delete(folder.serverRelativeUrl);
      return next;
    });
  }

  onDropOnFolder(
    event: CdkDragDrop<FolderNode | null, unknown, FileItem | FolderNode>,
    targetFolder: FolderNode,
  ): void {
    this.activeDropFolderUrl.set(null);
    const dragged = event.item.data;

    if (dragged.isFolder) {
      this.moveFolder(dragged, targetFolder);
      return;
    }

    this.moveFile(dragged, targetFolder);
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
    if (source.serverRelativeUrl === target.serverRelativeUrl) {
      return false;
    }
    if (this.isDescendant(target.serverRelativeUrl, source.serverRelativeUrl)) {
      return false;
    }

    return this.getParentFolderUrl(source.serverRelativeUrl) !== target.serverRelativeUrl;
  }

  private isDescendant(candidateUrl: string, ancestorUrl: string): boolean {
    let current = this.allFolders().find((folder) => folder.serverRelativeUrl === candidateUrl) ?? null;

    while (current) {
      const parentFolderUrl = this.getParentFolderUrl(current.serverRelativeUrl);
      if (parentFolderUrl === ancestorUrl) {
        return true;
      }
      current = parentFolderUrl
        ? (this.allFolders().find((folder) => folder.serverRelativeUrl === parentFolderUrl) ?? null)
        : null;
    }

    return false;
  }

  private canDropIntoFolder(dragged: FileItem | FolderNode, targetFolder: FolderNode | null): boolean {
    if (!targetFolder || this.isOperating()) {
      return false;
    }

    if (dragged.isFolder) {
      return this.isValidFolderMove(dragged, targetFolder);
    }

    return this.getParentFolderUrl(dragged.serverRelativeUrl) !== targetFolder.serverRelativeUrl;
  }

  private getParentFolderUrl(serverRelativeUrl: string): string | null {
    const lastSlashIndex = serverRelativeUrl.lastIndexOf('/');
    if (lastSlashIndex <= 0) {
      return null;
    }

    const parentFolderUrl = serverRelativeUrl.slice(0, lastSlashIndex);
    return parentFolderUrl.length > 0 ? parentFolderUrl : null;
  }

  private moveFile(file: FileItem, targetFolder: FolderNode): void {
    if (this.getParentFolderUrl(file.serverRelativeUrl) === targetFolder.serverRelativeUrl) {
      return;
    }

    const destinationServerRelativeUrl = this.buildDestinationUrl(targetFolder.serverRelativeUrl, file.name);
    this.pendingOperations.update((count) => count + 1);
    this.explorerService
      .moveFileTo(file.serverRelativeUrl, destinationServerRelativeUrl)
      .pipe(finalize(() => this.pendingOperations.update((count) => Math.max(0, count - 1))))
      .subscribe();
  }

  private moveFolder(folder: FolderNode, targetFolder: FolderNode): void {
    if (!this.isValidFolderMove(folder, targetFolder)) {
      return;
    }

    const destinationServerRelativeUrl = this.buildDestinationUrl(targetFolder.serverRelativeUrl, folder.name);
    this.pendingOperations.update((count) => count + 1);
    this.explorerService
      .moveFolderTo(folder.serverRelativeUrl, destinationServerRelativeUrl)
      .pipe(finalize(() => this.pendingOperations.update((count) => Math.max(0, count - 1))))
      .subscribe();
  }

  private selectFolder(folderUrl: string): void {
    this.selectedFolderUrl.set(folderUrl);
  }

  private buildDestinationUrl(parentFolderUrl: string, itemName: string): string {
    return `${parentFolderUrl}/${itemName}`;
  }

  private toTreeNode(
    folder: FolderNode,
    level: number,
    expandedFolderUrls: Set<string>,
    allFolders: FolderNode[],
  ): TreeNode<FolderNode> {
    const childFolders = allFolders.filter(
      (childFolder) => this.getParentFolderUrl(childFolder.serverRelativeUrl) === folder.serverRelativeUrl,
    );
    const isExpanded = expandedFolderUrls.has(folder.serverRelativeUrl);
    const isLoaded = this.explorerService.isFolderLoaded(folder.serverRelativeUrl);

    return {
      key: folder.serverRelativeUrl,
      label: folder.name,
      data: { ...folder, level },
      expanded: isExpanded,
      leaf: isLoaded && childFolders.length === 0,
      children: isExpanded
        ? childFolders.map((childFolder) => this.toTreeNode(childFolder, level + 1, expandedFolderUrls, allFolders))
        : [],
    };
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
