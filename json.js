const TelegramBot = require('node-telegram-bot-api');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');

// Токен вашего бота и канал для проверки
const token = '7251480208:AAHUy1k1qW__Gi6WlrixsHebwrDWFcYauoQ';
const requiredChannel = '@naneironkah'; 
const channelUrl = `https://t.me/${requiredChannel.replace('@', '')}`; // Ссылка на канал
const TelegramBot = require('node-telegram-bot-api');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const axios = require('axios');

// Токен вашего бота и канал для проверки
const token = '7251480208:AAHUy1k1qW__Gi6WlrixsHebwrDWFcYauoQ';
const requiredChannel = '@naneironkah'; 
const channelUrl = `https://t.me/${requiredChannel.replace('@', '')}`; // Ссылка на канал

// Создание экземпляра бота
const bot = new TelegramBot(token, { polling: true });

// Функция проверки подписки
async function isUserSubscribed(userId) {
  try {
    const response = await axios.get(`https://api.telegram.org/bot${token}/getChatMember`, {
      params: {
        chat_id: requiredChannel,
        user_id: userId,
      },
    });

    const status = response.data.result.status;
    return ['member', 'administrator', 'creator'].includes(status);
  } catch (error) {
    console.error('Ошибка проверки подписки:', error.response?.data || error.message);
    return false;
  }
}

// Извлечение текста из сложной структуры JSON
function extractText(message) {
  if (!message.text) return '';
  if (Array.isArray(message.text)) {
    return message.text.map((item) => (typeof item === 'string' ? item : item.text || '')).join('');
  }
  return message.text;
}

// Увеличение и уменьшение ширины колонок в XLSX
function adjustColumnWidths(worksheet) {
    const cols = Object.keys(worksheet).filter((key) => key[0] !== '!');
    const maxWidths = {};
  
    cols.forEach((cell) => {
      const col = cell.match(/[A-Z]+/)[0];
      const value = worksheet[cell]?.v?.toString() || '';
      maxWidths[col] = Math.max(maxWidths[col] || 0, value.length);
    });
  
    worksheet['!cols'] = Object.keys(maxWidths).map((col, index) => {
      // Уменьшить ширину для колонок "date" и "from", увеличить для остальных
      if (index === 0 || index === 1) {
        return { wch: Math.max(1, maxWidths[col] / 1) }; // Уменьшение в 9 раз
      }
      return { wch: Math.max(1, maxWidths[col] / 10) }; // Увеличение в 10 раз
    });
  }

// Обработчик команды /start
bot.onText(/\/start/, (msg) => {
  const chatId = msg.chat.id;

  // Отправляем сообщение с приветствием и кнопками
  bot.sendMessage(chatId, `Привет! 👋
Я помогу очистить JSON от не нужных символов и знаков 📂✨
Важно: бот работает только с файлами JSON!

Результат ты сможешь скачать в .txt или .xlsx

Чтобы я заработал, подпишись на канал ${requiredChannel}. Бот всегда бесплатный, но без подписки не запустится.

Подписался? Тогда загружай файл! 🚀 .`, {
    reply_markup: {
      inline_keyboard: [
        [{ text: 'Подписаться', url: channelUrl }]
      ],
    },
  });
});

// Обработчик callback-кнопок
bot.on('callback_query', async (callbackQuery) => {
  const msg = callbackQuery.message;
  const chatId = msg.chat.id;
  const userId = callbackQuery.from.id;
  const data = callbackQuery.data;

  if (data === 'check_subscription') {
    const subscribed = await isUserSubscribed(userId);

    if (subscribed) {
      bot.sendMessage(chatId, 'Вы успешно подписаны на канал. Теперь вы можете отправить JSON файл.');
    } else {
      bot.sendMessage(chatId, `Вы не подписаны на канал ${requiredChannel}. Пожалуйста, подпишитесь.`, {
        reply_markup: {
          inline_keyboard: [
            [{ text: 'Подписаться', url: channelUrl }],
            [{ text: 'Проверить подписку', callback_data: 'check_subscription' }],
          ],
        },
      });
    }

    bot.answerCallbackQuery(callbackQuery.id); // Закрываем обработку
  }
});

// Обработчик приема файла от пользователя
bot.on('document', async (msg) => {
  const chatId = msg.chat.id;
  const userId = msg.from.id;

  // Проверка подписки на канал
  if (!(await isUserSubscribed(userId))) {
    bot.sendMessage(chatId, `Ой-ой! 😅
Похоже, ты ещё не подписан на канал, а это обязательное условие для работы бота.`, {
      reply_markup: {
        inline_keyboard: [
          [{ text: 'Подписаться', url: channelUrl }],
          [{ text: 'Проверить подписку', callback_data: 'check_subscription' }],
        ],
      },
    });
    return;
  }

  const fileId = msg.document.file_id;

  // Запрос на получение информации о файле
  bot.getFile(fileId).then((file) => {
    // Скачивание файла
    bot.downloadFile(file.file_id, './').then((filePath) => {
      // Проверка на JSON файл
      if (path.extname(filePath) !== '.json') {
        bot.sendMessage(chatId, 'Пожалуйста, отправьте JSON файл.');
        return;
      }

      // Чтение и анализ файла
      fs.readFile(filePath, 'utf-8', (err, data) => {
        if (err) {
          bot.sendMessage(chatId, 'Ошибка при чтении файла.');
          return console.error(err);
        }

        try {
          const jsonData = JSON.parse(data);

          if (jsonData && jsonData.messages && jsonData.messages.length > 0) {
            // Фильтрация нужных данных
            const extractedMessages = jsonData.messages.map((message) => ({
              Дата: message.date || '',
              Пользователь: message.from || '',
              Текст: extractText(message),
            }));

            const xlsxFileName = `сообщения_${Date.now()}.xlsx`;
            const txtFileName = `сообщения_${Date.now()}.txt`;

            // Создание XLSX файла
            const workbook = xlsx.utils.book_new();
            const worksheet = xlsx.utils.json_to_sheet(extractedMessages);
            adjustColumnWidths(worksheet); // Настройка ширины колонок
            xlsx.utils.book_append_sheet(workbook, worksheet, 'Сообщения');
            xlsx.writeFile(workbook, xlsxFileName);

            // Создание TXT файла
            const txtContent = extractedMessages
              .map((msg) => `Дата: ${msg.Дата}\nПользователь: ${msg.Пользователь}\nТекст: ${msg.Текст}\n\n`)
              .join('');
            fs.writeFileSync(txtFileName, txtContent);

            // Отправка кнопок для выбора формата файла
            const buttons = {
              reply_markup: {
                inline_keyboard: [
                  [
                    { text: 'Скачать TXT', callback_data: `txt_${txtFileName}` },
                    { text: 'Скачать XLSX', callback_data: `xlsx_${xlsxFileName}` },
                  ],
                ],
              },
            };

            bot.sendMessage(chatId, 'Выберите формат файла для скачивания:', buttons);
          } else {
            bot.sendMessage(chatId, 'Файл не содержит текстовых сообщений.');
          }
        } catch (parseError) {
          bot.sendMessage(chatId, 'Ошибка при парсинге JSON.');
          console.error(parseError);
        }
      });
    });
  });
});

// Обработчик инлайн-кнопок для загрузки файлов
bot.on('callback_query', (callbackQuery) => {
  const msg = callbackQuery.message;
  const chatId = msg.chat.id;
  const data = callbackQuery.data;

  if (data.startsWith('txt_')) {
    const fileName = data.slice(4);
    bot.sendDocument(chatId, fileName).then(() => {
      fs.unlink(fileName, (err) => {
        if (err) console.error(err);
      });
    });
  } else if (data.startsWith('xlsx_')) {
    const fileName = data.slice(5);
    bot.sendDocument(chatId, fileName).then(() => {
      fs.unlink(fileName, (err) => {
        if (err) console.error(err);
      });
    });
  }

  bot.answerCallbackQuery(callbackQuery.id);
});

// Обработчик ошибок
bot.on('polling_error', (error) => {
  console.error(error);
});
// Создание экземпляра бота
const bot = new TelegramBot(token, { polling: true });

// Функция проверки подписки
async function isUserSubscribed(userId) {
  try {
    const response = await axios.get(`https://api.telegram.org/bot${token}/getChatMember`, {
      params: {
        chat_id: requiredChannel,
        user_id: userId,
      },
    });

    const status = response.data.result.status;
    return ['member', 'administrator', 'creator'].includes(status);
  } catch (error) {
    console.error('Ошибка проверки подписки:', error.response?.data || error.message);
    return false;
  }
}

// Извлечение текста из сложной структуры JSON
function extractText(message) {
  if (!message.text) return '';
  if (Array.isArray(message.text)) {
    return message.text.map((item) => (typeof item === 'string' ? item : item.text || '')).join('');
  }
  return message.text;
}

// Увеличение и уменьшение ширины колонок в XLSX
function adjustColumnWidths(worksheet) {
    const cols = Object.keys(worksheet).filter((key) => key[0] !== '!');
    const maxWidths = {};
  
    cols.forEach((cell) => {
      const col = cell.match(/[A-Z]+/)[0];
      const value = worksheet[cell]?.v?.toString() || '';
      maxWidths[col] = Math.max(maxWidths[col] || 0, value.length);
    });
  
    worksheet['!cols'] = Object.keys(maxWidths).map((col, index) => {
      // Уменьшить ширину для колонок "date" и "from", увеличить для остальных
      if (index === 0 || index === 1) {
        return { wch: Math.max(1, maxWidths[col] / 1) }; // Уменьшение в 9 раз
      }
      return { wch: Math.max(1, maxWidths[col] / 10) }; // Увеличение в 10 раз
    });
  }

// Обработчик команды /start
bot.onText(/\/start/, (msg) => {
  const chatId = msg.chat.id;

  // Отправляем сообщение с приветствием и кнопками
  bot.sendMessage(chatId, `Привет! 👋
Я помогу очистить JSON от не нужных символов и знаков 📂✨
Важно: бот работает только с файлами JSON!

Результат ты сможешь скачать в .txt или .xlsx

Чтобы я заработал, подпишись на канал ${requiredChannel}. Бот всегда бесплатный, но без подписки не запустится.

Подписался? Тогда загружай файл! 🚀 .`, {
    reply_markup: {
      inline_keyboard: [
        [{ text: 'Подписаться', url: channelUrl }]
      ],
    },
  });
});

// Обработчик callback-кнопок
bot.on('callback_query', async (callbackQuery) => {
  const msg = callbackQuery.message;
  const chatId = msg.chat.id;
  const userId = callbackQuery.from.id;
  const data = callbackQuery.data;

  if (data === 'check_subscription') {
    const subscribed = await isUserSubscribed(userId);

    if (subscribed) {
      bot.sendMessage(chatId, 'Вы успешно подписаны на канал. Теперь вы можете отправить JSON файл.');
    } else {
      bot.sendMessage(chatId, `Вы не подписаны на канал ${requiredChannel}. Пожалуйста, подпишитесь.`, {
        reply_markup: {
          inline_keyboard: [
            [{ text: 'Подписаться', url: channelUrl }],
            [{ text: 'Проверить подписку', callback_data: 'check_subscription' }],
          ],
        },
      });
    }

    bot.answerCallbackQuery(callbackQuery.id); // Закрываем обработку
  }
});

// Обработчик приема файла от пользователя
bot.on('document', async (msg) => {
  const chatId = msg.chat.id;
  const userId = msg.from.id;

  // Проверка подписки на канал
  if (!(await isUserSubscribed(userId))) {
    bot.sendMessage(chatId, `Ой-ой! 😅
Похоже, ты ещё не подписан на канал, а это обязательное условие для работы бота.`, {
      reply_markup: {
        inline_keyboard: [
          [{ text: 'Подписаться', url: channelUrl }],
          [{ text: 'Проверить подписку', callback_data: 'check_subscription' }],
        ],
      },
    });
    return;
  }

  const fileId = msg.document.file_id;

  // Запрос на получение информации о файле
  bot.getFile(fileId).then((file) => {
    // Скачивание файла
    bot.downloadFile(file.file_id, './').then((filePath) => {
      // Проверка на JSON файл
      if (path.extname(filePath) !== '.json') {
        bot.sendMessage(chatId, 'Пожалуйста, отправьте JSON файл.');
        return;
      }

      // Чтение и анализ файла
      fs.readFile(filePath, 'utf-8', (err, data) => {
        if (err) {
          bot.sendMessage(chatId, 'Ошибка при чтении файла.');
          return console.error(err);
        }

        try {
          const jsonData = JSON.parse(data);

          if (jsonData && jsonData.messages && jsonData.messages.length > 0) {
            // Фильтрация нужных данных
            const extractedMessages = jsonData.messages.map((message) => ({
              Дата: message.date || '',
              Пользователь: message.from || '',
              Текст: extractText(message),
            }));

            const xlsxFileName = `сообщения_${Date.now()}.xlsx`;
            const txtFileName = `сообщения_${Date.now()}.txt`;

            // Создание XLSX файла
            const workbook = xlsx.utils.book_new();
            const worksheet = xlsx.utils.json_to_sheet(extractedMessages);
            adjustColumnWidths(worksheet); // Настройка ширины колонок
            xlsx.utils.book_append_sheet(workbook, worksheet, 'Сообщения');
            xlsx.writeFile(workbook, xlsxFileName);

            // Создание TXT файла
            const txtContent = extractedMessages
              .map((msg) => `Дата: ${msg.Дата}\nПользователь: ${msg.Пользователь}\nТекст: ${msg.Текст}\n\n`)
              .join('');
            fs.writeFileSync(txtFileName, txtContent);

            // Отправка кнопок для выбора формата файла
            const buttons = {
              reply_markup: {
                inline_keyboard: [
                  [
                    { text: 'Скачать TXT', callback_data: `txt_${txtFileName}` },
                    { text: 'Скачать XLSX', callback_data: `xlsx_${xlsxFileName}` },
                  ],
                ],
              },
            };

            bot.sendMessage(chatId, 'Выберите формат файла для скачивания:', buttons);
          } else {
            bot.sendMessage(chatId, 'Файл не содержит текстовых сообщений.');
          }
        } catch (parseError) {
          bot.sendMessage(chatId, 'Ошибка при парсинге JSON.');
          console.error(parseError);
        }
      });
    });
  });
});

// Обработчик инлайн-кнопок для загрузки файлов
bot.on('callback_query', (callbackQuery) => {
  const msg = callbackQuery.message;
  const chatId = msg.chat.id;
  const data = callbackQuery.data;

  if (data.startsWith('txt_')) {
    const fileName = data.slice(4);
    bot.sendDocument(chatId, fileName).then(() => {
      fs.unlink(fileName, (err) => {
        if (err) console.error(err);
      });
    });
  } else if (data.startsWith('xlsx_')) {
    const fileName = data.slice(5);
    bot.sendDocument(chatId, fileName).then(() => {
      fs.unlink(fileName, (err) => {
        if (err) console.error(err);
      });
    });
  }

  bot.answerCallbackQuery(callbackQuery.id);
});

// Обработчик ошибок
bot.on('polling_error', (error) => {
  console.error(error);
});