import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { ExcelService } from './excel/excel.service';
import { ExcelController } from './excel/excel.controller';
import { ExcelModule } from './excel/excel.module';
import {TotalController} from './excel/total.controller'
@Module({
  imports: [ExcelModule],
  controllers: [AppController, ExcelController,TotalController],
  providers: [AppService, ExcelService],
})
export class AppModule {}
