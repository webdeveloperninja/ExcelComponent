import { Injectable } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
import { Client } from '@microsoft/microsoft-graph-client';

import { User } from '@microsoft/microsoft-graph-types';
import { environment } from 'src/environments/environment';

@Injectable({
  providedIn: 'root'
})
export class AuthService {
  public authenticated: boolean;
  public user: User;

  constructor(private msalService: MsalService) {
    this.authenticated = this.msalService.getUser() != null;
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

    const graphClient = Client.init({
      authProvider: async done => {
        const token = await this.getAccessToken().catch(reason => {
          done(reason, null);
        });

        if (token) {
          done(null, token);
        } else {
          done('Could not get an access token', null);
        }
      }
    });

    // Get the user from Graph (GET /me)
    const graphUser = (await graphClient.api('/me').get()) as User;

    return graphUser;
  }
}
