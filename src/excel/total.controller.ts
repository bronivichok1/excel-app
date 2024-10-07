import { Controller, Get, Res } from '@nestjs/common';
import { Response } from 'express';
import * as XLSX from 'xlsx';

@Controller('/total')
export class TotalController {
    @Get()
    async getTotal(@Res() res: Response) {
        try {
            // Загрузка книги
            const workbook = XLSX.readFile('Zhurnal.xlsx');
            const totalWorksheet = workbook.Sheets['Общее количество ']; // Убедитесь, что имя листа точно

            const results: number[] = [];

            // Извлечение значений из ячеек C4:C25
            for (let i = 4; i <= 25; i++) {
                const cellAddress = `C${i}`;
                const cellValue = totalWorksheet[cellAddress]?.v; // Получаем значение ячейки
                if (cellValue !== undefined && cellValue !== null) {
                    results.push(cellValue);
                }
            }

            // Возвращаем результат
            console.log(results)
            return res.status(200).json({ results });

        } catch (error) {
            console.error(error);
            return res.status(500).json({ error: 'Произошла ошибка при обработке файла.' });
        }
    }
}
