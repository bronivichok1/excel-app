import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config'; 
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { ExcelService } from './excel/excel.service';
import { ExcelController } from './excel/excel.controller';
import { ExcelModule } from './excel/excel.module';
import { TotalController } from './excel/total.controller';
import { TotalController2 } from './excel/total2.controller';
import { AllDataController } from './excel/alldata.controller';
import { Edit } from './excel/edit.controller';
import { DownloadController } from './excel/download.controller';
import { AuthModule } from './excel/auth.module'; 
import {EditClockController} from './excel/editclock.controller';

@Module({
  imports: [
    ConfigModule.forRoot({
      isGlobal: true, 
      envFilePath: '.env', 
    }),
    ExcelModule, 
    AuthModule
  ], 
  controllers: [
    AppController, 
    ExcelController, 
    TotalController,
    TotalController2, 
    AllDataController, 
    Edit, 
    DownloadController,
    EditClockController
  ],
  providers: [AppService, ExcelService],
})
export class AppModule {}
