import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { ExcelsheetComponent } from './excelsheet/excelsheet.component';
import { ExcelService } from './excel.service';

@NgModule({
  declarations: [AppComponent, ExcelsheetComponent],
  imports: [BrowserModule, AppRoutingModule],
  providers: [ExcelService],
  bootstrap: [AppComponent],
})
export class AppModule {}
