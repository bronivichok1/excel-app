import { Controller, Post, Body } from '@nestjs/common';
import { ExcelService } from './excel.service';

@Controller('excel')
export class ExcelController {
  constructor(private readonly excelService: ExcelService) {}

  @Post('add')
  async addData(@Body() data: { name: string; age: number }) {
    await this.excelService.addDataToExcel(data);
    return { message: 'Данные были добавлены в файл Excel.' };
  }
}