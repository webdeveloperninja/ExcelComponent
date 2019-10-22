import { Component, OnInit } from '@angular/core';
import { WorkbookWorksheet, WorkbookTable, WorkbookTableRow, RemoteItem, BaseItem } from '@microsoft/microsoft-graph-types';
import { GraphService } from '../../graph.service';
import { AuthService } from '../../auth.service';

@Component({
  selector: 'app-excel',
  templateUrl: './excel-client.component.html',
  styleUrls: ['./excel-client.component.scss']
})
export class ExcelComponent implements OnInit {
  readonly isAuthenticated = this.authService.authenticated;

  isLoading = false;

  workSheets: WorkbookWorksheet[];
  selectedWorksheet: WorkbookWorksheet;
  tables: WorkbookTable[];
  selectedTable: WorkbookTable;
  rows: WorkbookTableRow[];
  driveItems: BaseItem[];
  selectedDriveItem: BaseItem;

  get webUrl() {
    return this.selectedDriveItem.webUrl;
  }

  get eTagGuid() {
    return this.selectedDriveItem.eTag.match(/\{(.*?)\}/)[0];
  }

  get iframeUrl() {
    // tslint:disable-next-line:max-line-length
    return `https://robertdeveloper-my.sharepoint.com/personal/robert_robertdeveloper_onmicrosoft_com/_layouts/15/Doc.aspx?sourcedoc=${this.eTagGuid}&action=embedview&Item=${this.selectedTable.name}&wdDownloadButton=True&wdInConfigurator=True`;
  }

  constructor(private graphService: GraphService, private readonly authService: AuthService) {}

  async onWorkbookSelection(item: RemoteItem) {
    this.isLoading = true;

    this.selectedDriveItem = item;
    this.selectedWorksheet = null;
    this.workSheets = null;
    this.selectedTable = null;
    this.tables = null;
    this.workSheets = await this.graphService.getWorksheets(this.selectedDriveItem.name);

    this.isLoading = false;
  }

  async onTableSelection(table: WorkbookTable) {
    this.isLoading = true;

    this.selectedTable = table;
    this.rows = await this.graphService.getRows(this.selectedDriveItem.name, this.selectedWorksheet.name, this.selectedTable.name);

    this.isLoading = false;
  }

  async onWorksheetSelection(worksheet: WorkbookWorksheet) {
    this.isLoading = true;

    this.selectedWorksheet = worksheet;
    this.selectedTable = null;
    this.tables = null;
    this.tables = await this.graphService.getTables(this.selectedDriveItem.name, this.selectedWorksheet.name);

    this.isLoading = false;
  }

  async ngOnInit() {
    this.isLoading = true;

    const driveItems = await this.graphService.getDriveItems();
    const excelBooks = driveItems.filter(item => item.name.endsWith('xlsx'));
    this.driveItems = excelBooks;

    this.isLoading = false;
  }

  async signIn() {
    await this.authService.signIn();
  }
}
