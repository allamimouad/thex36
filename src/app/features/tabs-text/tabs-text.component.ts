import { ChangeDetectionStrategy, Component } from '@angular/core';

@Component({
  selector: 'app-tabs-text',
  standalone: true,
  changeDetection: ChangeDetectionStrategy.OnPush,
  templateUrl: './tabs-text.component.html',
  styleUrl: './tabs-text.component.scss',
})
export class TabsTextComponent {}
