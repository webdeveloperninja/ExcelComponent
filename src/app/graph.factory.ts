import { Injectable } from '@angular/core';
import { Client } from '@microsoft/microsoft-graph-client';
import { environment } from 'src/environments/environment';
import { MsalService } from '@azure/msal-angular';

@Injectable({
  providedIn: 'root'
})
export class GraphFactory {
  constructor(private readonly msalService: MsalService) {}

  create() {
    return Client.init({
      authProvider: async done => {
        const token = await this.msalService.acquireTokenSilent(environment.scopes);

        if (token) {
          done(null, token);
        } else {
          done('Could not get an access token', null);
        }
      }
    });
  }
}
