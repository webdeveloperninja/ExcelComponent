import { Component, OnInit } from '@angular/core';
import { WorkbookWorksheet, WorkbookTable, WorkbookTableRow } from '@microsoft/microsoft-graph-types';
import { GraphService } from '../../graph.service';
import { AuthService } from '../../auth.service';

@Component({
  selector: 'app-excel',
  templateUrl: './excel-client.component.html',
  styleUrls: ['./excel-client.component.scss']
})
export class ExcelComponent implements OnInit {
  readonly workbookName = 'Book.xlsx';

  isLoading = false;

  readonly isAuthenticated = this.authService.authenticated;
  workSheets: WorkbookWorksheet[];
  selectedWorksheet: WorkbookWorksheet;

  tables: WorkbookTable[];
  selectedTable: WorkbookTable;

  rows: WorkbookTableRow[];

  constructor(private graphService: GraphService, private readonly authService: AuthService) {}

  async onTableSelection(table: WorkbookTable) {
    this.isLoading = true;
    this.selectedTable = table;

    this.rows = await this.graphService.getRows(this.workbookName, this.selectedWorksheet.name, this.selectedTable.name);

    this.isLoading = false;
  }

  async onWorksheetSelection(worksheet: WorkbookWorksheet) {
    this.isLoading = true;

    this.selectedWorksheet = worksheet;
    this.tables = await this.graphService.getTables(this.workbookName, this.selectedWorksheet.name);

    this.isLoading = false;
  }

  async ngOnInit() {
    this.isLoading = true;
    this.workSheets = await this.graphService.getWorksheets(this.workbookName);
    this.isLoading = false;
  }

  async signIn() {
    await this.authService.signIn();
  }
}
