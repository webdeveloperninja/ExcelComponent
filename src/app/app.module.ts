import { DoBootstrap, Injector, NgModule } from '@angular/core';
import { createCustomElement } from '@angular/elements';
import { BrowserModule } from '@angular/platform-browser';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { RouterModule } from '@angular/router';
import { MsalModule } from '@azure/msal-angular';
import { library } from '@fortawesome/fontawesome-svg-core';
import { faUserCircle } from '@fortawesome/free-regular-svg-icons';
import { faExternalLinkAlt } from '@fortawesome/free-solid-svg-icons';
import { environment } from 'src/environments/environment';
import { ExcelComponent } from './components/excel-client/excel-client.component';
import { ThemeModule } from './theme.module';

library.add(faExternalLinkAlt);
library.add(faUserCircle);

@NgModule({
  declarations: [ExcelComponent],
  imports: [
    RouterModule.forRoot([]),
    ThemeModule,
    BrowserModule,
    MsalModule.forRoot({
      clientID: environment.appId
    }),
    BrowserAnimationsModule
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
