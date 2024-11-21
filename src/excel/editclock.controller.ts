import { Controller, Get, Query, Res } from '@nestjs/common';
import { Response } from 'express';
import * as XLSX from 'xlsx';

@Controller('editclock')
export class EditClockController {
  
  @Get()
  async getRowData(@Query('rowNumber') rowNumber: string, @Res() res: Response) {
    try {
      const rowNumberInt = parseInt(rowNumber, 10);
      if (isNaN(rowNumberInt)) {
        return res.status(400).json({ error: 'Некорректный номер строки. Убедитесь, что это число.' });
      }

      const workbook = XLSX.readFile('Zhurnal.xlsx');
      const worksheet = workbook.Sheets['Учет актов'];

      if (!worksheet) {
        return res.status(404).json({ error: 'Лист "Учет актов" не найден в файле.' });
      }

      const effectiveRowNumber = rowNumberInt + 5;
      const results = [];
      const columnCount = 100; 

      for (let col = 16; col < 16 + columnCount; col += 4) {
        const rowData = [];

        for (let i = 0; i < 4; i++) {
          const cellAddress = XLSX.utils.encode_cell({ r: effectiveRowNumber - 1, c: col + i });
          const cellValue = worksheet[cellAddress] ? worksheet[cellAddress].v : null;
          rowData.push(cellValue);
        }

        results.push(rowData);
      }

      if (results.length === 0) {
        return res.status(404).json({ error: 'Нет данных для данной записи.' });
      }

      return res.json(results);

    } catch (error) {
      console.error('Ошибка при чтении файла Excel:', error);
      return res.status(500).send('Не удалось прочитать файл Excel.');
    }
  }
}
