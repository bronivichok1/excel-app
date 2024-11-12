import { Controller, Post, Body, Res } from '@nestjs/common';
import { Response } from 'express';
import * as XLSX from 'xlsx';

@Controller('total')
export class TotalController {
    @Post()
    async getTotal(@Body() body: { number: string }, @Res() res: Response) {
        // Преобразуем строку в число
        const number = parseInt(body.number, 10); 
        
        // Проверяем, является ли преобразованное значение числом
        if (isNaN(number)) {
            return res.status(400).json({ error: 'Недопустимый номер записи.' });
        }

        try {
            // Загружаем рабочую книгу из Excel файла
            const workbook = XLSX.readFile('Zhurnal.xlsx');
            const totalWorksheet = workbook.Sheets['Общий список ']; 
            const accountWorksheet = workbook.Sheets['Учет актов'];

            const results: any[] = []; 

            // Вычисляем номер строки: 107 + number
            const rowNumber = 107 + number; 

            // Получаем данные из столбцов B до U
            for (let col = 1; col <= 20; col++) { // Столбцы B (1) до U (20)
                const cellAddress = XLSX.utils.encode_cell({ r: rowNumber - 1, c: col });
                const cellValue = totalWorksheet[cellAddress] ? totalWorksheet[cellAddress].v : null;

                // Если ячейка не пустая, добавляем значение в результирующий массив
                if (cellValue !== null) {
                    results.push(cellValue);
                }
            }

            // Проверяем, есть ли какие-либо результаты из первого листа
            if (results.length === 0) {
                return res.status(404).json({ error: 'Нет данных для данной записи.' });
            }

            // Получаем значения столбцов O и P с листа "Учет актов" на строке number + 5
            const accountRowNumber = number + 5;
            const valuesFromAccountSheet = [];

            for (let colIndex = 14; colIndex <= 15; colIndex++) { // Столбцы O (14) и P (15)
                const cellAddress = XLSX.utils.encode_cell({ r: accountRowNumber - 1, c: colIndex });
                const cellValue = accountWorksheet[cellAddress] ? accountWorksheet[cellAddress].v : null;

                // Если ячейка не пустая, добавляем значение в массив
                if (cellValue !== null) {
                    valuesFromAccountSheet.push(cellValue);
                }
            }
            results.push(...valuesFromAccountSheet);

            
            return res.json(results);

        } catch (error) {
            console.error('Ошибка при чтении файла Excel:', error);
            return res.status(500).send('Не удалось прочитать файл Excel.');
        }
    }
}
