import { DoBootstrap, Injector, NgModule } from '@angular/core';
import { createCustomElement } from '@angular/elements';
import { BrowserModule } from '@angular/platform-browser';
import { MsalModule } from '@azure/msal-angular';
import { FontAwesomeModule } from '@fortawesome/angular-fontawesome';
import { library } from '@fortawesome/fontawesome-svg-core';
import { faUserCircle } from '@fortawesome/free-regular-svg-icons';
import { faExternalLinkAlt } from '@fortawesome/free-solid-svg-icons';
import { NgbModule } from '@ng-bootstrap/ng-bootstrap';
import { ExcelComponent } from './components/excel-client/excel-client.component';
import { RouterModule } from '@angular/router';
import { environment } from 'src/environments/environment';

library.add(faExternalLinkAlt);
library.add(faUserCircle);

@NgModule({
  declarations: [ExcelComponent],
  imports: [
    RouterModule.forRoot([]),
    BrowserModule,
    NgbModule,
    FontAwesomeModule,
    MsalModule.forRoot({
      clientID: environment.appId
    })
  ],
  entryComponents: [ExcelComponent]
})
export class AppModule implements DoBootstrap {
  constructor(private injector: Injector) {}

  ngDoBootstrap() {
    const excel = createCustomElement(ExcelComponent, { injector: this.injector });
    customElements.define('excel-client', excel);
  }
}
