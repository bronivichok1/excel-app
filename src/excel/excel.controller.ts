import { Controller, Post, Body } from '@nestjs/common';
import { ExcelService } from './excel.service';
import { CreateDataDto } from './create-data.dto';

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
    prim:string;
  }) {
    await this.excelService.addDataToExcel(data);
    return { message: 'Данные были добавлены в файл Excel.' };
  }
  @Post('red')
  async redData(@Body() data: {
    number:number;
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
    prim:string;
  }) {
    await this.excelService.redDataToExcel(data);
    return { message: 'Данные были добавлены в файл Excel.' };
  }
  @Post('clock')
    async clock(@Body() data: CreateDataDto) {
        await this.excelService.clockToExcel(data);
        return { message: 'Данные были добавлены в файл Excel.' };
    }
}
