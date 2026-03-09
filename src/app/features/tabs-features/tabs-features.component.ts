import { ChangeDetectionStrategy, Component } from '@angular/core';
import { TabsModule } from 'primeng/tabs';
import { SharepointExplorerComponent } from '../sharepoint-explorer/sharepoint-explorer.component';
import { TabsTextComponent } from '../tabs-text/tabs-text.component';

@Component({
  selector: 'app-tabs-features',
  standalone: true,
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [TabsModule, SharepointExplorerComponent, TabsTextComponent],
  templateUrl: './tabs-features.component.html',
  styleUrl: './tabs-features.component.scss',
})
export class TabsFeaturesComponent {}
