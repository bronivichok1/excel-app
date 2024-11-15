import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config'; 
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { ExcelService } from './excel/excel.service';
import { ExcelController } from './excel/excel.controller';
import { ExcelModule } from './excel/excel.module';
import { TotalController } from './excel/total.controller';
import { AllDataController } from './excel/alldata.controller';
import { Edit } from './excel/edit.controller';
import { DownloadController } from './excel/download.controller';
import { AuthModule } from './excel/auth.module'; 

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
    AllDataController, 
    Edit, 
    DownloadController
  ],
  providers: [AppService, ExcelService],
})
export class AppModule {}
