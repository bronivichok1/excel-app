import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';

async function bootstrap() {
  const app = await NestFactory.create(AppModule);

  // Включение CORS
  app.enableCors({
    origin:'*',
    /*origin: 'http://localhost:3001', // Разрешите доступ только с этого источника
    methods: 'GET,HEAD,PUT,PATCH,POST,DELETE',
    allowedHeaders: 'Content-Type, Authorization',
    credentials: true, // Укажите true, если ваши запросы требуют учетных данных (например, куки)*/
  });

  await app.listen(3000);
}
bootstrap();
