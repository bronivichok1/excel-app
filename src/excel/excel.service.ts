import { Injectable, Logger } from '@nestjs/common';
import * as ExcelJS from 'exceljs';

@Injectable()
export class ExcelService {
  private readonly logger = new Logger(ExcelService.name);

  async updateExcelFile() {
    const workbook = new ExcelJS.Workbook();

    try {
      // Чтение существующего файла Excel
      await workbook.xlsx.readFile('yourfile.xlsx'); // Замените на название вашего файла

      const worksheet = workbook.getWorksheet('Sheet1'); // Укажите название листа

      // Внесение данных в определенные строки и колонки
      worksheet.getCell('A1').value = 'Новый заголовок'; // Внести новый заголовок в ячейку A1
      worksheet.getCell('B2').value = 'Иван'; // Внести имя в ячейку B2
      worksheet.getCell('C2').value = 30; // Внести возраст в ячейку C2

      // Сохранение изменений
      await workbook.xlsx.writeFile('updatedfile.xlsx'); // Сохранение в новый файл или перезапись существующего
      this.logger.log('Файл Excel обновлён');
    } catch (error) {
      this.logger.error('Ошибка при обновлении файла:', error);
    }
  }
}
