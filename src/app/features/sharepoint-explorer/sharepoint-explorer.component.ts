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
import { ExplorerRow, FileItem, FolderNode } from './models/sharepoint-explorer.models';
import {
  SHAREPOINT_EXPLORER_ROOT_URL,
  SharepointExplorerService,
} from './services/sharepoint-explorer.service';

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
  readonly rootFolder = computed<FolderNode | null>(() => {
    return {
      isFolder: true,
      name: '/',
      serverRelativeUrl: SHAREPOINT_EXPLORER_ROOT_URL,
      level: 0,
    };
  });
  readonly selectedFolderUrl = signal<string>('');
  readonly activeDropFolderUrl = signal<string | null>(null);
  readonly activeNativeUploadFolderUrl = signal<string | null>(null);
  readonly expandedFolderUrls = signal<Set<string>>(new Set<string>());
  readonly nativeUploadSummary = signal<string | null>(null);
  readonly pendingOperations = signal(0);
  readonly isOperating = computed(() => this.pendingOperations() > 0);

  readonly treeNodes = computed<TreeNode<FolderNode>[]>(() => {
    const rootFolder = this.rootFolder();
    if (!rootFolder) {
      return [];
    }

    const expandedFolderUrls = this.expandedFolderUrls();
    const childFolders = this.rootFolders();
    const isExpanded = expandedFolderUrls.has(rootFolder.serverRelativeUrl);

    return [
      {
        key: rootFolder.serverRelativeUrl,
        label: rootFolder.name,
        data: rootFolder,
        expanded: isExpanded,
        leaf: false,
        children: isExpanded
          ? childFolders.map((childFolder) => this.toTreeNode(childFolder, 1, expandedFolderUrls, this.allFolders()))
          : [],
      },
    ];
  });
  readonly selectedItems = toSignal(
    toObservable(this.selectedFolderUrl).pipe(
      switchMap((folderUrl) =>
        folderUrl ? this.explorerService.getFolderContentOf(folderUrl) : of([] as ExplorerRow[])
      ),
    ),
    { initialValue: [] as ExplorerRow[] },
  );
  readonly breadcrumb = computed(() => this.formatSelectedPath(this.selectedFolderUrl()));
  readonly selectedFolder = computed(
    () =>
      this.allFolders().find((folder) => folder.serverRelativeUrl === this.selectedFolderUrl()) ??
      (this.selectedFolderUrl() === this.rootFolder()?.serverRelativeUrl ? this.rootFolder() : null),
  );

  constructor() {
    this.explorerService
      .loadRootFolders()
      .pipe(takeUntilDestroyed(this.destroyRef))
      .subscribe();

    effect(() => {
      const rootFolder = this.rootFolder();
      if (!rootFolder || this.selectedFolderUrl()) {
        return;
      }

      this.selectedFolderUrl.set(rootFolder.serverRelativeUrl);
      this.expandedFolderUrls.set(new Set([rootFolder.serverRelativeUrl]));
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

  onNativeUploadDragOver(event: DragEvent, targetFolderUrl: string): void {
    if (!this.hasNativeFiles(event)) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    this.activeNativeUploadFolderUrl.set(targetFolderUrl);

    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = 'copy';
    }
  }

  onNativeUploadDragLeave(event: DragEvent, targetFolderUrl: string): void {
    if (!this.hasNativeFiles(event)) {
      return;
    }

    const relatedTarget = event.relatedTarget;
    if (relatedTarget instanceof Node && event.currentTarget instanceof Node && event.currentTarget.contains(relatedTarget)) {
      return;
    }

    if (this.activeNativeUploadFolderUrl() === targetFolderUrl) {
      this.activeNativeUploadFolderUrl.set(null);
    }
  }

  onNativeUploadDrop(event: DragEvent, targetFolder: FolderNode): void {
    if (!this.hasNativeFiles(event)) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    this.activeNativeUploadFolderUrl.set(null);

    const files = this.explorerService.extractUploadFiles(event);
    if (files.length === 0) {
      return;
    }

    this.pendingOperations.update((count) => count + 1);
    this.explorerService
      .uploadFilesTo(targetFolder.serverRelativeUrl, files)
      .pipe(finalize(() => this.pendingOperations.update((count) => Math.max(0, count - 1))))
      .subscribe();
    this.nativeUploadSummary.set(this.explorerService.buildUploadSummary(targetFolder.serverRelativeUrl, files));
  }

  onTreeNodeExpand(event: { node: TreeNode<FolderNode> }): void {
    const folder = event.node.data;
    if (!folder) {
      return;
    }

    this.expandedFolderUrls.update((current) => new Set(current).add(folder.serverRelativeUrl));
    if (folder.serverRelativeUrl === this.rootFolder()?.serverRelativeUrl) {
      return;
    }

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

  private formatSelectedPath(folderUrl: string): string {
    if (!folderUrl || folderUrl === SHAREPOINT_EXPLORER_ROOT_URL) {
      return '/';
    }

    const relativePath = folderUrl.startsWith(SHAREPOINT_EXPLORER_ROOT_URL)
      ? folderUrl.slice(SHAREPOINT_EXPLORER_ROOT_URL.length)
      : folderUrl;

    if (!relativePath || relativePath === '/') {
      return '/';
    }

    const normalizedPath = relativePath.startsWith('/') ? relativePath : `/${relativePath}`;
    return normalizedPath.replaceAll('/', ' / ');
  }

  private hasNativeFiles(event: DragEvent): boolean {
    const dataTransfer = event.dataTransfer;
    return !!dataTransfer && Array.from(dataTransfer.types).includes('Files');
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
