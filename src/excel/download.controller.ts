import { Controller, Get, Res } from '@nestjs/common';
import { Response } from 'express';
import * as fs from 'fs';
import * as path from 'path';

@Controller('download')
export class DownloadController {
  @Get('zhurnal')
  downloadExcel(@Res() res: Response) {
    const filePath = path.join(__dirname, '..', '../Zhurnal.xlsx'); 
    
   
    fs.access(filePath, fs.constants.F_OK, (err) => {
      if (err) {
        return res.status(404).send('File not found');
      }

      res.setHeader('Content-Disposition', 'attachment; filename=Zhurnal.xlsx');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      
      const fileStream = fs.createReadStream(filePath);
      fileStream.pipe(res);
    });
  }
}
