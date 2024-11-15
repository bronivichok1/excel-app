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

        try {

            const workbook = XLSX.readFile('Zhurnal.xlsx');
            const totalWorksheet = workbook.Sheets['Общий список ']; 
            const accountWorksheet = workbook.Sheets['Учет актов'];

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
            const valuesFromAccountSheet = [];
            
            for (let colIndex = 14; colIndex <= 15; colIndex++) { 
                const cellAddress = XLSX.utils.encode_cell({ r: accountRowNumber - 1, c: colIndex });
                let cellValue = null;
            
                try {
                    const json = XLSX.utils.sheet_to_json(accountWorksheet, { header: 1, range: cellAddress + ':' + cellAddress });
                    cellValue = json[0][0];
            
                } catch (error) {
                    console.error("Ошибка при обработке ячейки:", cellAddress, error);
                }
            
                if (cellValue !== null && cellValue !== undefined) {
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
