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
  }) {
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    // Пробуем прочитать существующий файл Excel
    try {
      await workbook.xlsx.readFile('Zhurnal.xlsx');
      worksheet = workbook.getWorksheet('Общий список ');
      // Если лист отсутствует, создаем новый
    } catch (error) {
      this.logger.error('Ошибка при чтении файла: ', error);
      throw new Error('Не удалось прочитать файл Excel.');
    }

    // Поиск первой пустой строки, начиная с B108
    let rowIndex = 108;
    const maxRows = worksheet.rowCount; // Получаем количество строк в листе

    while (rowIndex <= maxRows) {
      const row = worksheet.getRow(rowIndex);
      const isEmpty = !row.getCell(2).value && !row.getCell(3).value && !row.getCell(4).value;

      if (isEmpty) {
        break;
      }
      rowIndex++;
    }

    // Записываем данные в соответствующие ячейки
    const targetRow = worksheet.getRow(rowIndex);
    targetRow.getCell(2).value = data.surname; // B
    targetRow.getCell(3).value = data.name; // C
    targetRow.getCell(4).value = data.othername; // D
    targetRow.getCell(5).value = data.kafedra; // E
    targetRow.getCell(6).value = data.VO; // F
    targetRow.getCell(7).value = data.DOV; // G
    targetRow.getCell(8).value = data.workplace; // H
    targetRow.getCell(9).value = data.orgcategory; // I
    targetRow.getCell(10).value = data.worktitlecategory; // J
    targetRow.getCell(11).value = data.studystep; // K
    targetRow.getCell(12).value = data.studyrang; // L
    targetRow.getCell(13).value = data.kvalcategory; // M
    targetRow.getCell(14).value = data.oldstatus; // N
    targetRow.getCell(15).value = data.olddata; // O
    targetRow.getCell(16).value = data.datanotification; // P
    targetRow.getCell(17).value = data.numberdoc; // Q
    targetRow.getCell(18).value = data.numberdocdop; // R

    await this.updateFormulas(workbook, ['Учет актов', 'Списки по кафедрам', 'Общее количество ']); // Укажите названия листов

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
