import { Controller, Get, Res } from '@nestjs/common';
import { Response } from 'express';
import * as XLSX from 'xlsx';

@Controller('/edit')
export class Edit {
    @Get()
    async getTotal(@Res() res: Response) {
        try {
            const workbook = XLSX.readFile('Zhurnal.xlsx');
            const totalWorksheet = workbook.Sheets['Общее количество ']; 

            const results: number[] = [];

            for (let i = 4; i <= 25; i++) {
                const cellAddress = `C${i}`;
                const cellValue = totalWorksheet[cellAddress]?.v; 
                if (cellValue !== undefined && cellValue !== null) {
                    results.push(cellValue);
                }
            }

            return res.status(200).json({ results });

        } catch (error) {
            console.error(error);
            return res.status(500).json({ error: 'Произошла ошибка при обработке файла.' });
        }
    }
}
