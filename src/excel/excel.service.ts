import { Injectable, Logger } from '@nestjs/common';
import * as ExcelJS from 'exceljs';

@Injectable()
export class ExcelService {
  private readonly logger = new Logger(ExcelService.name);

  async addDataToExcel(data: {
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
    prim: string;
  }) {
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    try {
      await workbook.xlsx.readFile('Zhurnal.xlsx');
      worksheet = workbook.getWorksheet('Общий список ');
    } catch (error) {
      this.logger.error('Ошибка при чтении файла: ', error);
      throw new Error('Не удалось прочитать файл Excel.');
    }

   
    let rowIndex = 108;
    const maxRows = worksheet.rowCount; 

    while (rowIndex <= maxRows) {
      const row = worksheet.getRow(rowIndex);
      const isEmpty = !row.getCell(2).value && !row.getCell(3).value && !row.getCell(4).value;

      if (isEmpty) {
        break;
      }
      rowIndex++;
    }

   
    const targetRow = worksheet.getRow(rowIndex);
    targetRow.getCell(2).value = data.surname; 
    targetRow.getCell(3).value = data.name; 
    targetRow.getCell(4).value = data.othername; 
    targetRow.getCell(5).value = data.kafedra; 
    targetRow.getCell(6).value = data.VO; 
    targetRow.getCell(7).value = data.DOV; 
    targetRow.getCell(10).value = data.workplace; 
    targetRow.getCell(11).value = data.orgcategory; 
    targetRow.getCell(12).value = data.worktitlecategory; 
    targetRow.getCell(13).value = data.studystep; 
    targetRow.getCell(14).value = data.studyrang; 
    targetRow.getCell(15).value = data.kvalcategory; 
    targetRow.getCell(16).value = data.oldstatus; 
    targetRow.getCell(17).value = data.olddata; 
    targetRow.getCell(18).value = data.datanotification; 
    targetRow.getCell(19).value = data.numberdoc; 
    targetRow.getCell(20).value = data.prim; 
    targetRow.getCell(21).value = data.numberdocdop; 

    await this.updateFormulas(workbook, ['Учет актов', 'Списки по кафедрам', 'Общее количество ']); 

    // Сохранение изменений в файл
    await workbook.xlsx.writeFile('Zhurnal.xlsx');
    this.logger.log(`Данные добавлены в строку ${rowIndex}: ${JSON.stringify(data)}`);
  }
  private async updateFormulas(workbook: ExcelJS.Workbook, sheetNames: string[]) {
    for (const sheetName of sheetNames) {
      const worksheet = workbook.getWorksheet(sheetName);
      if (worksheet) {
        // Обходим все ячейки, чтобы инициировать пересчет формул
        worksheet.eachRow((row) => {
          row.eachCell((cell) => {
            if (cell.formula) {
              // Перезаписываем текущие значения, чтобы вызвать пересчет
              const originalValue = cell.value;
              cell.value = null; // Установить временно значение null
              cell.value = originalValue; // Обратно устанавливаем оригинальное значение
            }
          });
        });
      }
    }
  }
}
