import { Injectable } from '@angular/core';
import { Client } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { AuthService } from './auth.service';

@Injectable({
  providedIn: 'root'
})
export class GraphService {
  private graphClient: Client;
  constructor(private authService: AuthService) {
    this.graphClient = Client.init({
      authProvider: async done => {
        const token = await this.authService.getAccessToken().catch(reason => {
          done(reason, null);
        });

        if (token) {
          done(null, token);
        } else {
          done('Could not get an access token', null);
        }
      }
    });
  }

  async getDriveItems(): Promise<MicrosoftGraph.RemoteItem[]> {
    const result = await this.graphClient.api(`https://graph.microsoft.com/v1.0/me/drive/root/children`).get();

    return result.value;
  }

  async getWorksheets(workBookName: string): Promise<MicrosoftGraph.WorkbookWorksheet[]> {
    const result = await this.graphClient.api(`/me/drive/root:/${workBookName}:/workbook/worksheets`).get();

    return result.value;
  }

  async getTables(workBookName: string, workSheetName: string): Promise<MicrosoftGraph.WorkbookTable[]> {
    const result = await this.graphClient.api(`/me/drive/root:/${workBookName}:/workbook/worksheets/${workSheetName}/tables`).get();

    return result.value;
  }

  async getRows(workBookName: string, workSheetName: string, tableName: string): Promise<MicrosoftGraph.WorkbookTableRow[]> {
    const result = await this.graphClient
      .api(`/me/drive/root:/${workBookName}:/workbook/worksheets/${workSheetName}/tables/${tableName}/rows`)
      .get();

    return result.value;
  }
}
