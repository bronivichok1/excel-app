import { Injectable, Logger } from '@nestjs/common';
import * as ExcelJS from 'exceljs';
import { CreateDataDto } from './create-data.dto'; 
import * as XLSX from 'xlsx';

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
    prim: string;
  }) {
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    try {
      await workbook.xlsx.readFile('Zhurnal.xlsx');
      worksheet = workbook.getWorksheet('Общий список ');
    } catch (error) {
      this.logger.error('Ошибка при чтении файла: ', error);
      throw new Error('Не удалось прочитать файл Excel.');
    }

   
    let rowIndex = 108;
    const maxRows = worksheet.rowCount; 

    while (rowIndex <= maxRows) {
      const row = worksheet.getRow(rowIndex);
      const isEmpty = !row.getCell(2).value && !row.getCell(3).value && !row.getCell(4).value;

      if (isEmpty) {
        break;
      }
      rowIndex++;
    }

   
    const targetRow = worksheet.getRow(rowIndex);
    targetRow.getCell(2).value = data.surname; 
    targetRow.getCell(3).value = data.name; 
    targetRow.getCell(4).value = data.othername; 
    targetRow.getCell(5).value = data.kafedra; 
    targetRow.getCell(6).value = data.VO; 
    targetRow.getCell(7).value = data.DOV; 
    targetRow.getCell(10).value = data.workplace; 
    targetRow.getCell(11).value = data.orgcategory; 
    targetRow.getCell(12).value = data.worktitlecategory; 
    targetRow.getCell(13).value = data.studystep; 
    targetRow.getCell(14).value = data.studyrang; 
    targetRow.getCell(15).value = data.kvalcategory; 
    targetRow.getCell(16).value = data.oldstatus; 
    targetRow.getCell(17).value = data.olddata; 
    targetRow.getCell(18).value = data.datanotification; 
    targetRow.getCell(19).value = data.numberdoc; 
    targetRow.getCell(20).value = data.prim; 
    targetRow.getCell(21).value = data.numberdocdop; 

    await this.updateFormulas(workbook, ['Учет актов', 'Списки по кафедрам', 'Общее количество ']); 

    await workbook.xlsx.writeFile('Zhurnal.xlsx');
    this.logger.log(`Данные добавлены в строку ${rowIndex}: ${JSON.stringify(data)}`);
  }
  private async updateFormulas(workbook: ExcelJS.Workbook, sheetNames: string[]) {
    for (const sheetName of sheetNames) {
      const worksheet = workbook.getWorksheet(sheetName);
      if (worksheet) {
        worksheet.eachRow((row) => {
          row.eachCell((cell) => {
            if (cell.formula) {
              const originalValue = cell.value;
              cell.value = null; 
              cell.value = originalValue; 
            }
          });
        });
      }
    }
  }

  async redDataToExcel(data: {
    number: number; 
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
    prim: string;
}) {
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    try {
        await workbook.xlsx.readFile('Zhurnal.xlsx');
        worksheet = workbook.getWorksheet('Общий список ');
    } catch (error) {
        this.logger.error('Ошибка при чтении файла: ', error);
        throw new Error('Не удалось прочитать файл Excel.');
    }

    const rowIndex = data.number + 107;
    const row = worksheet.getRow(rowIndex);
    
    const isSame = 
        row.getCell(2).value === data.surname &&
        row.getCell(3).value === data.name &&
        row.getCell(4).value === data.othername &&
        row.getCell(5).value === data.kafedra &&
        row.getCell(10).value === data.workplace &&
        row.getCell(11).value === data.orgcategory &&
        row.getCell(12).value === data.worktitlecategory &&
        row.getCell(13).value === data.studystep &&
        row.getCell(14).value === data.studyrang &&
        row.getCell(15).value === data.kvalcategory &&
        row.getCell(16).value === data.oldstatus &&
        row.getCell(17).value === data.olddata &&
        row.getCell(18).value === data.datanotification &&
        row.getCell(19).value === data.numberdoc &&
        row.getCell(20).value === data.prim &&
        row.getCell(21).value === data.numberdocdop;

    if (isSame) {
        this.logger.log(`Данные уже существуют в строке ${rowIndex}, обновление не требуется.`);
        return; 
    }

    row.getCell(2).value = data.surname; 
    row.getCell(3).value = data.name; 
    row.getCell(4).value = data.othername; 
    row.getCell(5).value = data.kafedra; 
    row.getCell(6).value = data.VO; 
    row.getCell(7).value = data.DOV; 
    row.getCell(10).value = data.workplace; 
    row.getCell(11).value = data.orgcategory; 
    row.getCell(12).value = data.worktitlecategory; 
    row.getCell(13).value = data.studystep; 
    row.getCell(14).value = data.studyrang; 
    row.getCell(15).value = data.kvalcategory; 
    row.getCell(16).value = data.oldstatus; 
    row.getCell(17).value = data.olddata; 
    row.getCell(18).value = data.datanotification; 
    row.getCell(19).value = data.numberdoc; 
    row.getCell(20).value = data.prim; 
    row.getCell(21).value = data.numberdocdop; 

    await this.updateFormulas(workbook, ['Учет актов', 'Списки по кафедрам', 'Общее количество ']); 

    await workbook.xlsx.writeFile('Zhurnal.xlsx');
    this.logger.log(`Данные добавлены в строку ${rowIndex}: ${JSON.stringify(data)}`);
  }

  async clockToExcel(data: CreateDataDto) {
    const { number, additionalFields } = data; 
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    try {
        await workbook.xlsx.readFile('Zhurnal.xlsx');
        worksheet = workbook.getWorksheet('Учет актов');
    } catch (error) {
        console.error('Ошибка при чтении файла: ', error); 
        throw new Error('Не удалось прочитать файл Excel.');
    }

    const startRowIndex = number + 5 ; 
    const fields = additionalFields;
    
    const updateCell = (cellReference: string, value: any, isNumber: boolean) => {
      if (value !== undefined && value !== null && value !== '') {
        if (isNumber) {
          const numericValue = parseFloat(value);
          if (!isNaN(numericValue)) {
            worksheet.getCell(cellReference).value = numericValue;
          }
        } else {
          worksheet.getCell(cellReference).value = value;
        }
      }
    };

  updateCell(`Q${startRowIndex}`, fields[0].date, false);      
  updateCell(`R${startRowIndex}`, fields[0].month, false);     
  updateCell(`S${startRowIndex}`, fields[0].hoursVO, true);   
  updateCell(`T${startRowIndex}`, fields[0].hoursDOV, true);  
  
  updateCell(`U${startRowIndex}`, fields[1].date, false);      
  updateCell(`V${startRowIndex}`, fields[1].month, false);     
  updateCell(`W${startRowIndex}`, fields[1].hoursVO, true);   
  updateCell(`X${startRowIndex}`, fields[1].hoursDOV, true);  
  
  updateCell(`Y${startRowIndex}`, fields[2].date, false);      
  updateCell(`Z${startRowIndex}`, fields[2].month, false);     
  updateCell(`AA${startRowIndex}`,fields[2].hoursVO, true);   
  updateCell(`AB${startRowIndex}`, fields[2].hoursDOV, true);  
  
  updateCell(`AC${startRowIndex}`, fields[3].date, false);      
  updateCell(`AD${startRowIndex}`, fields[3].month, false);     
  updateCell(`AE${startRowIndex}`, fields[3].hoursVO, true);   
  updateCell(`AF${startRowIndex}`, fields[3].hoursDOV, true);  
  
  updateCell(`AG${startRowIndex}`, fields[4].date, false);      
  updateCell(`AH${startRowIndex}`, fields[4].month, false);     
  updateCell(`AI${startRowIndex}`, fields[4].hoursVO, true);   
  updateCell(`AJ${startRowIndex}`, fields[4].hoursDOV, true);  
  
  updateCell(`AK${startRowIndex}`, fields[5].date, false);      
  updateCell(`AL${startRowIndex}`, fields[5].month, false);     
  updateCell(`AM${startRowIndex}`, fields[5].hoursVO, true);   
  updateCell(`AN${startRowIndex}`, fields[5].hoursDOV, true);  
  
  updateCell(`AO${startRowIndex}`, fields[6].date, false);      
  updateCell(`AP${startRowIndex}`, fields[6].month, false);     
  updateCell(`AQ${startRowIndex}`, fields[6].hoursVO, true);   
  updateCell(`AR${startRowIndex}`, fields[6].hoursDOV, true);  
  
  updateCell(`AS${startRowIndex}`, fields[7].date, false);      
  updateCell(`AT${startRowIndex}`, fields[7].month, false);     
  updateCell(`AU${startRowIndex}`, fields[7].hoursVO, true);   
  updateCell(`AV${startRowIndex}`, fields[7].hoursDOV, true);  
  
  updateCell(`AW${startRowIndex}`, fields[8].date, false);      
  updateCell(`AX${startRowIndex}`, fields[8].month, false);     
  updateCell(`AY${startRowIndex}`, fields[8].hoursVO, true);   
  updateCell(`AZ${startRowIndex}`, fields[8].hoursDOV, true);  
  
  updateCell(`BA${startRowIndex}`, fields[9].date, false);      
  updateCell(`BB${startRowIndex}`, fields[9].month, false);     
  updateCell(`BC${startRowIndex}`, fields[9].hoursVO, true);   
  updateCell(`BD${startRowIndex}`, fields[9].hoursDOV, true);  
  
  updateCell(`BE${startRowIndex}`, fields[10].date, false);      
  updateCell(`BF${startRowIndex}`, fields[10].month, false);     
  updateCell(`BG${startRowIndex}`, fields[10].hoursVO, true);   
  updateCell(`BH${startRowIndex}`, fields[10].hoursDOV, true);  
  
  updateCell(`BI${startRowIndex}`, fields[11].date, false);      
  updateCell(`BJ${startRowIndex}`, fields[11].month, false);     
  updateCell(`BK${startRowIndex}`, fields[11].hoursVO, true);   
  updateCell(`BL${startRowIndex}`, fields[11].hoursDOV, true);  
  
  updateCell(`BM${startRowIndex}`, fields[12].date, false);      
  updateCell(`BN${startRowIndex}`, fields[12].month, false);     
  updateCell(`BO${startRowIndex}`, fields[12].hoursVO, true);   
  updateCell(`BP${startRowIndex}`, fields[12].hoursDOV, true);  
  
  updateCell(`BQ${startRowIndex}`, fields[13].date, false);      
updateCell(`BR${startRowIndex}`, fields[13].month, false);     
updateCell(`BS${startRowIndex}`, fields[13].hoursVO, true);   
updateCell(`BT${startRowIndex}`, fields[13].hoursDOV, true);  

updateCell(`BU${startRowIndex}`, fields[14].date, false);      
updateCell(`BV${startRowIndex}`, fields[14].month, false);     
updateCell(`BW${startRowIndex}`, fields[14].hoursVO, true);   
updateCell(`BX${startRowIndex}`, fields[14].hoursDOV, true);  

updateCell(`BY${startRowIndex}`, fields[15].date, false);      
updateCell(`BZ${startRowIndex}`, fields[15].month, false);     
updateCell(`CA${startRowIndex}`, fields[15].hoursVO, true);   
updateCell(`CB${startRowIndex}`, fields[15].hoursDOV, true);  

updateCell(`CC${startRowIndex}`, fields[16].date, false);      
updateCell(`CD${startRowIndex}`, fields[16].month, false);     
updateCell(`CE${startRowIndex}`, fields[16].hoursVO, true);   
updateCell(`CF${startRowIndex}`, fields[16].hoursDOV, true);  

updateCell(`CG${startRowIndex}`, fields[17].date, false);      
updateCell(`CH${startRowIndex}`, fields[17].month, false);     
updateCell(`CI${startRowIndex}`, fields[17].hoursVO, true);   
updateCell(`CJ${startRowIndex}`, fields[17].hoursDOV, true);  

updateCell(`CK${startRowIndex}`, fields[18].date, false);      
updateCell(`CL${startRowIndex}`, fields[18].month, false);     
updateCell(`CM${startRowIndex}`, fields[18].hoursVO, true);   
updateCell(`CN${startRowIndex}`, fields[18].hoursDOV, true);  

updateCell(`CO${startRowIndex}`, fields[19].date, false);      
updateCell(`CP${startRowIndex}`, fields[19].month, false);     
updateCell(`CQ${startRowIndex}`, fields[19].hoursVO, true);   
updateCell(`CR${startRowIndex}`, fields[19].hoursDOV, true);  

updateCell(`CS${startRowIndex}`, fields[20].date, false);      
updateCell(`CT${startRowIndex}`, fields[20].month, false);     
updateCell(`CU${startRowIndex}`, fields[20].hoursVO, true);   
updateCell(`CV${startRowIndex}`, fields[20].hoursDOV, true);  

updateCell(`CW${startRowIndex}`, fields[21].date, false);      
updateCell(`CX${startRowIndex}`, fields[21].month, false);     
updateCell(`CY${startRowIndex}`, fields[21].hoursVO, true);   
updateCell(`CZ${startRowIndex}`, fields[21].hoursDOV, true);  

updateCell(`DA${startRowIndex}`, fields[22].date, false);      
updateCell(`DB${startRowIndex}`, fields[22].month, false);     
updateCell(`DC${startRowIndex}`, fields[22].hoursVO, true);   
updateCell(`DD${startRowIndex}`, fields[22].hoursDOV, true);  

updateCell(`DE${startRowIndex}`, fields[23].date, false);      
updateCell(`DF${startRowIndex}`, fields[23].month, false);     
updateCell(`DG${startRowIndex}`, fields[23].hoursVO, true);   
updateCell(`DH${startRowIndex}`, fields[23].hoursDOV, true);  

updateCell(`DI${startRowIndex}`, fields[24].date, false);      
updateCell(`DJ${startRowIndex}`, fields[24].month, false);     
updateCell(`DK${startRowIndex}`, fields[24].hoursVO, true);   
updateCell(`DL${startRowIndex}`, fields[24].hoursDOV, true);   

  

  await this.updateFormulas(workbook, ['Учет актов', 'Списки по кафедрам', 'Общее количество ']);
    try {
        await workbook.xlsx.writeFile('Zhurnal.xlsx'); 
    } catch (error) {
        console.error('Ошибка при записи файла: ', error); 
        throw new Error('Не удалось записать файл Excel.');
    }
}
}
  



