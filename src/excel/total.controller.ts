import { Controller, Post, Body, Res } from '@nestjs/common';
import { Response } from 'express';
import * as XLSX from 'xlsx';

@Controller('total')
export class TotalController {
    @Post()
    async getTotal(@Body() body: { number: string }, @Res() res: Response) {
        const number = parseInt(body.number, 10);

        if (isNaN(number)) {
            return res.status(400).json({ error: 'Недопустимый номер записи.' });
        }

        let workbook;
        try {
            workbook = XLSX.readFile('Zhurnal.xlsx');
        } catch (error) {
            console.error('Ошибка при чтении файла Excel:', error);
            return res.status(500).send('Не удалось прочитать файл Excel.');
        }

        const totalWorksheet = workbook.Sheets['Общий список '];
        const accountWorksheet = workbook.Sheets['Учет актов'];

        await this.updateFormulas(workbook, ['Учет актов', 'Общий список ', 'Общее количество ']);

        const results: any[] = [];
        const rowNumber = 107 + number; 

        for (let col = 1; col <= 20; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: rowNumber - 1, c: col });
            const cellValue = totalWorksheet[cellAddress] ? totalWorksheet[cellAddress].v : null;

            if (cellValue !== null) {
                results.push(cellValue); 
            }
        }

        if (results.length === 0) {
            return res.status(404).json({ error: 'Нет данных для данной записи.' });
        }

        const accountRowNumber = number + 5;

        // Получаем значения O и P, используя заданные формулы
        const sumO = this.calculateStaticSum(accountWorksheet, accountRowNumber, [
            'S', 'W', 'AA', 'AE', 'AI', 'AM', 'AQ', 'AU', 'AY', 
            'BC', 'BG', 'BK', 'BO', 'BS', 'BW', 'CA', 'CE', 'CI', 
            'CM', 'CQ', 'CU', 'CY', 'DC', 'DG', 'DK'
        ]);

        const sumP = this.calculateStaticSum(accountWorksheet, accountRowNumber, [
            'T', 'X', 'AB', 'AF', 'AJ', 'AN', 'AR', 'AV', 'AZ', 
            'BD', 'BH', 'BL', 'BP', 'BT', 'BX', 'CB', 'CF', 'CJ', 
            'CN', 'CR', 'CV', 'CZ', 'DD', 'DH', 'DL'
        ]);

        // Преобразуем значения sumO и sumP в текстовый формат
        const sumOText = sumO.toString();
        const sumPText = sumP.toString();

        // Добавляем вычисленные суммы в результаты
        results.push(sumOText, sumPText);

        console.log('Результаты:', results);

        return res.json(results); // Возврат результатов
    }

    private calculateStaticSum(worksheet: XLSX.WorkSheet, row: number, columns: string[]): number {
        let total = 0;

        for (const column of columns) {
            const cellAddress = column + row; 
            const cell = worksheet[cellAddress];
            if (cell && cell.v !== undefined) {
                total += typeof cell.v === 'number' ? cell.v : 0; 
            }
        }

        return total;
    }
    private async updateFormulas(workbook: XLSX.WorkBook, sheetNames: string[]) {
        for (const sheetName of sheetNames) {
            const worksheet = workbook.Sheets[sheetName];
            if (worksheet) {
                const range = XLSX.utils.decode_range(worksheet['!ref']);
                for (let r = range.s.r; r <= range.e.r; r++) {
                    for (let c = range.s.c; c <= range.e.c; c++) {
                        const cellAddress = XLSX.utils.encode_cell({ r, c });
                        const cell = worksheet[cellAddress];
                        if (cell && cell.f) {
                            const formula = cell.f;
                            cell.f = formula; 
                            delete cell.t; 
                        }
                    }
                }
            }
        }
    }
}
