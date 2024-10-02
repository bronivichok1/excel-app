import { Controller, Post, Body } from '@nestjs/common';
import { ExcelService } from './excel.service';

@Controller('excel')
export class ExcelController {
  constructor(private readonly excelService: ExcelService) {}

  @Post('add')
  async addData(@Body() data: {
    surname: string;
    name: string;
    othername: string;
    kafedra: string;
    workplace: string;
    orgcategory: string;
    worktitlecategory: string;
    studyrang: string;
    studystep: string;
    kvalcategory: string;
    oldstatus: string;
    olddata: string;
    datanotification: string;
    numberdoc: string;
    numberdocdop: string;
    VO: string;
    DOV: string;
  }) {
    await this.excelService.addDataToExcel(data);
    return { message: 'Данные были добавлены в файл Excel.' };
  }
}
