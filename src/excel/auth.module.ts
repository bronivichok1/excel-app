import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config';
import { AuthController } from './status.controller'; 

@Module({
  imports: [ConfigModule], 
  controllers: [AuthController],
})
export class AuthModule {}
