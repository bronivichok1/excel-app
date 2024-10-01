import { Module } from '@nestjs/common';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { ExcelService } from './excel/excel.service';
import { ExcelController } from './excel/excel.controller';
import { ExcelModule } from './excel/excel.module';

@Module({
  imports: [ExcelModule],
  controllers: [AppController, ExcelController],
  providers: [AppService, ExcelService],
})
export class AppModule {}
