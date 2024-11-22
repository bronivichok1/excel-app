import { Controller, Get, Res } from '@nestjs/common';
import { Response } from 'express';
import * as XLSX from 'xlsx';

@Controller('/totaldata')
export class TotalController2 {
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
                }else{
                    results.push(0);
                }
            }

            return res.status(200).json({ results });

        } catch (error) {
            console.error(error);
            return res.status(500).json({ error: 'Произошла ошибка при обработке файла.' });
        }
        
    }
    
}