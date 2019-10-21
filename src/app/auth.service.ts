import { Injectable } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
import { Client } from '@microsoft/microsoft-graph-client';

import { User } from '@microsoft/microsoft-graph-types';
import { environment } from 'src/environments/environment';
import { GraphFactory } from './graph.factory';

@Injectable({
  providedIn: 'root'
})
export class AuthService {
  private readonly graphClient: Client;
  public authenticated: boolean;
  public user: User;

  constructor(private readonly msalService: MsalService, graphFactory: GraphFactory) {
    this.authenticated = this.msalService.getUser() != null;
    this.graphClient = graphFactory.create();

    this.getUser().then(user => {
      this.user = user;
    });
  }

  async signIn(): Promise<void> {
    const result = await this.msalService.loginPopup(environment.scopes).catch(reason => {});

    if (result) {
      this.authenticated = true;
    }
  }

  signOut(): void {
    this.msalService.logout();
    this.user = null;
    this.authenticated = false;
  }

  async getAccessToken(): Promise<string> {
    return await this.msalService.acquireTokenSilent(environment.scopes).catch(reason => {});
  }

  private async getUser(): Promise<User> {
    if (!this.authenticated) {
      return null;
    }

    return this.graphClient.api('/me').get();
  }
}
