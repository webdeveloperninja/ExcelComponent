import { Component, OnInit } from '@angular/core';
import { WorkbookWorksheet } from '@microsoft/microsoft-graph-types';
import { GraphService } from '../../graph.service';
import { AuthService } from '../../auth.service';

@Component({
  selector: 'app-excel',
  templateUrl: './excel-client.component.html'
})
export class ExcelComponent implements OnInit {
  readonly isAuthenticated = this.authService.authenticated;
  workSheets: WorkbookWorksheet[];

  constructor(private graphService: GraphService, private readonly authService: AuthService) {}

  async ngOnInit() {
    this.workSheets = await this.graphService.getWorksheets('Book.xlsx');
  }

  async signIn() {
    await this.authService.signIn();
  }
}
