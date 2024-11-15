import { Controller, Post, Body } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';

@Controller('auth') // Путь к ресурсу
export class AuthController {
  constructor(private configService: ConfigService) {}

  @Post()
  async authenticate(@Body() body: { login: string; password: string }): Promise<{ status: string; message?: string }> {
    const { login, password } = body;

    // Сравнение с переменными окружения
    if (login == this.configService.get<string>('LOGIN1') && password == this.configService.get<string>('PASSWORD1')) {
      return { status: 'red' };
    } else if (login == this.configService.get<string>('LOGIN2') && password == this.configService.get<string>('PASSWORD2')) {
      return { status: 'red' };
    } else if (login == this.configService.get<string>('LOGIN3') && password == this.configService.get<string>('PASSWORD3')) {
      return { status: 'watch' };
    } else {
      return { status: 'none', message: 'Неверные учетные данные' };
    }
  }
}
