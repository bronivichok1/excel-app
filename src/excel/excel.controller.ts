import { Controller, Get } from '@nestjs/common';
import { ExcelService } from './excel.service';

@Controller('excel')
export class ExcelController {
  constructor(private readonly excelService: ExcelService) {}

  @Get('update')
  async updateFile() {
    await this.excelService.updateExcelFile();
    return { message: 'Файл был обновлён.' };
  }
}
