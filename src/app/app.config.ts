import { ApplicationConfig, provideBrowserGlobalErrorListeners } from '@angular/core';
import { provideRouter } from '@angular/router';

import { routes } from './app.routes';
import { providePrimeNG } from 'primeng/config';
import Aura from '@primeuix/themes/aura';
import { SharepointExplorerClient } from './features/sharepoint-explorer/services/sharepoint-explorer.client';
import { SimulatedSharepointExplorerClientService } from './features/sharepoint-explorer/services/simulated-sharepoint-explorer-client.service';

export const appConfig: ApplicationConfig = {
  providers: [
    provideBrowserGlobalErrorListeners(),
    provideRouter(routes),
    {
      provide: SharepointExplorerClient,
      useClass: SimulatedSharepointExplorerClientService,
    },
    providePrimeNG({
      theme: {
        preset: Aura,
      },
    }),
  ],
};
