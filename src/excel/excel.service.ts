import { Injectable, Logger } from '@nestjs/common';
import * as ExcelJS from 'exceljs';

@Injectable()
export class ExcelService {
  private readonly logger = new Logger(ExcelService.name);

  async addDataToExcel(data: { name: string; age: number }) {
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    // Попробуем прочитать файл, если он существует
    try {
      await workbook.xlsx.readFile('Zhurnal.xlsx'); // Замените на название вашего файла
      worksheet = workbook.getWorksheet('Sheet1');
    } catch (error) {
      // Если файл не существует, создадим новый
      worksheet = workbook.addWorksheet('Sheet1');
      worksheet.columns = [
        { header: 'Имя', key: 'name', width: 20 },
        { header: 'Возраст', key: 'age', width: 10 },
      ];
    }

    // Добавление новой строки с данными
    worksheet.addRow({
      name: data.name,
      age: data.age,
    });

    // Сохранение изменений в файл
    await workbook.xlsx.writeFile('data.xlsx');
    this.logger.log(`Данные добавлены: ${JSON.stringify(data)}`);
  }
}
