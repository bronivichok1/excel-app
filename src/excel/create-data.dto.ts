import { IsNotEmpty, IsString, IsNumber, ValidateNested } from 'class-validator';
import { Type } from 'class-transformer';

class AdditionalField {
    @IsString()
    date: string;

    @IsString()
    month: string;

    @IsNumber()
    hoursVO: number;

    @IsNumber()
    hoursDOV: number;
}

export class CreateDataDto {
    @IsNumber() 
    number: number; 

    @IsString()
    VO: string;

    @IsString()
    DOV: string;

    @IsString()
    VOconst: string;

    @IsString()
    DOVconst: string;

    @ValidateNested({ each: true })
    @Type(() => AdditionalField)
    additionalFields: AdditionalField[];
}
