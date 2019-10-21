import { Injectable } from '@angular/core';
import { Client } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { GraphFactory } from './graph.factory';

@Injectable({
  providedIn: 'root'
})
export class GraphService {
  private graphClient: Client;

  constructor(graphFactory: GraphFactory) {
    this.graphClient = graphFactory.create();
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
