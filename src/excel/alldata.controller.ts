import { Controller, Get, Res } from '@nestjs/common';
import { Response } from 'express';
import * as XLSX from 'xlsx';

@Controller('alldata')
export class AllDataController {
    @Get()
    async getAllData(@Res() res: Response) {
        try {
            const workbook = XLSX.readFile('Zhurnal.xlsx');
            const worksheet = workbook.Sheets['Общий список ']; 

            const data: any[] = []; 

            const maxRow = 2648; 

            for (let row = 108; row <= maxRow; row++) {

                const cellA = worksheet[`A${row}`] ? worksheet[`A${row}`].v : null;
                const cellB = worksheet[`B${row}`] ? worksheet[`B${row}`].v : null;
                const cellC = worksheet[`C${row}`] ? worksheet[`C${row}`].v : null;
                const cellD = worksheet[`D${row}`] ? worksheet[`D${row}`].v : null;
                const cellE = worksheet[`E${row}`] ? worksheet[`E${row}`].v : null;


                if (cellA !== null && cellB !== null && cellC !== null && cellD !== null&& cellE !== null) {
                    data.push([cellA, cellB, cellC, cellD, cellE]);
                }
            }

            return res.json(data);
        } catch (error) {
            console.error('Ошибка при чтении файла Excel:', error);
            return res.status(500).send('Не удалось прочитать файл Excel.');
        }
    }
}
