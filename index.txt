const { Client, GatewayIntentBits, Routes } = require('discord.js');
const fetch = (...args) => import('node-fetch').then(({ default: f }) => f(...args));
const XLSX = require('xlsx');
const {
  EmbedBuilder,
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle
} = require('@discordjs/builders');
require('dotenv').config();

// ❗ Перевірка ENV-змінних
if (!process.env.SHEET_ID || !process.env.GOOGLE_API_KEY || !process.env.APP_SCRIPT_URL || !process.env.BOT_TOKEN) {
  console.error("❗ Одна з обов'язкових ENV-змінних відсутня");
  process.exit(1);
}

const SHEET_ID = process.env.SHEET_ID;
const SHEET_NAME = 'Аркуш1';
const GOOGLE_API_KEY = process.env.GOOGLE_API_KEY;
const APP_SCRIPT_URL = process.env.APP_SCRIPT_URL;

const API_URL = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${SHEET_NAME}?key=${GOOGLE_API_KEY}`;
const CELLS_URL = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/C964:E964?key=${GOOGLE_API_KEY}`;

const client = new Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMessages,
    GatewayIntentBits.MessageContent,
    GatewayIntentBits.GuildMessageReactions
  ]
});

// Реєстрація слеш-команд
const { REST } = require('@discordjs/rest');
const { version } = require('discord.js').Constants;

const commands = [
  {
    name: 'залишки',
    description: 'Показує загальну кількість товарів'
  },
  {
    name: 'оновити',
    description: 'Принудово оновлює дані'
  },
  {
    name: 'порожні',
    description: 'Показує товари з мінімальною кількістю'
  },
  {
    name: 'пошук',
    description: 'Пошук за полями таблиці',
    options: [
      {
        name: 'поле',
        description: 'За яким полем шукати',
        type: 3,
        required: true,
        choices: [
          { name: 'Найменування', value: 'назва' },
          { name: 'Серійний номер', value: 'серія' },
          { name: 'Контрагент', value: 'контрагент' },
          { name: 'Кількість', value: 'кількість' },
          { name: 'Ціна', value: 'ціна' }
        ]
      },
      {
        name: 'запит',
        description: 'Що шукати (наприклад: "стол", "5")',
        type: 3,
        required: true
      }
    ]
  },
  {
    name: 'розумний-пошук',
    description: 'Пошук за кількома полями',
    options: [
      { name: 'номенклатура', type: 3, description: 'Шукати за назвою товару' },
      { name: 'контрагент', type: 3, description: 'Шукати за контрагентом' },
      { name: 'серія', type: 3, description: 'Шукати за серійним номером' },
      { name: 'ціна_вище', type: 10, description: 'Показувати товари дорожче цього значення' },
      { name: 'кількість_вище', type: 10, description: 'Показувати товари з кількістю більше' }
    ]
  },
  {
    name: 'пошук-експортовано',
    description: 'Експортує результати пошуку в Excel (.xlsx)',
    options: [
      {
        name: 'поле',
        description: 'За яким полем шукати',
        type: 3,
        required: true,
        choices: [
          { name: 'Найменування', value: 'назва' },
          { name: 'Серійний номер', value: 'серія' },
          { name: 'Контрагент', value: 'контрагент' },
          { name: 'Кількість', value: 'кількість' },
          { name: 'Ціна', value: 'ціна' }
        ]
      },
      {
        name: 'запит',
        description: 'Що шукати (наприклад: "стол", "5")',
        type: 3,
        required: true
      }
    ]
  },
  {
    name: 'експорт',
    description: 'Експортує всю таблицю в Excel (.xlsx)'
  },
  {
    name: 'help',
    description: 'Показує список усіх доступних команд'
  }
];

// ────────────────
// 📁 Функція для отримання даних із Google Sheets
// ────────────────
async function getSheetData(range = SHEET_NAME) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/${range}?key=${GOOGLE_API_KEY}`;
  console.log(`🔍 Завантажую дані з ${url}`);

  try {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
    
    const data = await res.json();
    return data.values || [];
  } catch (err) {
    console.error('⚠️ Не вдалося отримати дані:', err.message);
    return [];
  }
}

const rest = new REST({ version: '10' }).setToken(process.env.BOT_TOKEN);

let previousData = null;

// Функція для завантаження даних з Google Sheets
async function loadSheetData() {
  const res = await fetch(API_URL);
  if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
  return await res.json();
}

// Функція для отримання індексу колонки за її назвою
function getColumnIndex(headers, field) {
  const headerMap = {
    назва: ['назва', 'найменування'],
    серія: ['серійний номер', 'серія'],
    контрагент: ['контрагент', 'постачальник'],
    кількість: ['кількість', 'залишок'],
    ціна: ['ціна', 'вартість']
  };

  for (let i = 0; i < headers.length; i++) {
    const headerName = headers[i]?.toLowerCase().trim();
    if (headerMap[field].includes(headerName)) {
      return i;
    }
  }

  return -1;
}

// Обробка змін у таблиці
async function checkForChanges(botClient) {
  try {
    const data = await loadSheetData();
    const currentRows = data.values;

    if (!previousData) {
      previousData = currentRows;
      return;
    }

    const changedCells = [];
    for (let i = 0; i < Math.min(currentRows.length, previousData.length); i++) {
      const oldRow = previousData[i];
      const newRow = currentRows[i];
      if (!oldRow || !newRow) continue;

      for (let j = 0; j < Math.min(oldRow.length, newRow.length); j++) {
        if (oldRow[j] !== newRow[j]) {
          changedCells.push({
            row: i + 1,
            column: j + 1,
            from: oldRow[j],
            to: newRow[j]
          });
        }
      }
    }

    if (changedCells.length > 0) {
      const channel = botClient.channels.cache.find(ch => ch.name === 'склад' && ch.type === 0);
      if (!channel) return;

      let message = '🔔 Виявлено зміни в таблиці:\n';
      changedCells.forEach(change => {
        const colLetter = String.fromCharCode(64 + change.column);
        message += `\nКлітинка ${colLetter}${change.row}:\nБуло: \`${change.from}\`, стало: \`${change.to}\``;
      });

      const embed = new EmbedBuilder()
        .setTitle('🔔 Виявлено зміни')
        .setDescription(message)
        .setColor(3447003)
        .setTimestamp();

      await channel.send({ embeds: [embed] });
    }

    previousData = currentRows;
  } catch (err) {
    console.error('❌ Не вдалося перевірити зміни:', err.message);
  }
}

// ────────────────
// 📊 Обробка слеш-команд
// ────────────────
client.on('interactionCreate', async interaction => {
  if (!interaction.isChatInputCommand()) return;
  try {
    switch (interaction.commandName) {
      case 'залишки':
        const cellRes = await fetch(CELLS_URL);
        if (!cellRes.ok) throw new Error(`HTTP error! status: ${cellRes.status}`);
        const cellData = await cellRes.json();
        const cellValues = cellData.values?.flat() || [];
        const totalValue = Number(cellValues[0]) || 0;
        const totalQuantity = Number(cellValues[1]) || 0;
        const avgPrice = Number(cellValues[2]) || 0;
        const embed = new EmbedBuilder()
          .setTitle('📊 Загальні залишки')
          .addFields([
            { name: 'Загальна сума', value: `${totalValue.toFixed(2)} грн`, inline: true },
            { name: 'Кількість', value: `${totalQuantity} шт.`, inline: true },
            { name: 'Середня ціна', value: `${avgPrice.toFixed(2)} грн`, inline: true }
          ])
          .setColor(5763719)
          .setFooter({ text: 'Фінансова служба' })
          .setTimestamp();
        await interaction.reply({ embeds: [embed], ephemeral: false });
        break;
      case 'оновити':
        const sheetData = await getSheetData();
        const rows = sheetData.slice(1);
        const headers = sheetData[0];
        let output = '| Назва       | Кількість | Ціна |\n';
        output += '|--------------|------------|--------|\n';
        for (let i = Math.max(0, rows.length - 10); i < rows.length; i++) {
          const row = rows[i];
          const name = row[getColumnIndex(headers, 'назва')] || '—';
          const quantity = row[getColumnIndex(headers, 'кількість')] || '—';
          const price = row[getColumnIndex(headers, 'ціна')] || '—';
          output += `| ${name.padEnd(13).slice(0, 13)} | ${quantity} | ${price} |\n`;
        }
        const embedUpdate = new EmbedBuilder()
          .setTitle('🔄 Останні записи')
          .setDescription(`\`\`\`md\n${output}\`\`\``)
          .setColor(3066993);
        await interaction.reply({ embeds: [embedUpdate], ephemeral: false });
        break;
      case 'порожні':
        const lowStockData = await getSheetData();
        const lowStockRows = lowStockData.slice(1);
        const lowStockHeaders = lowStockData[0];
        const lowStock = lowStockRows.filter(row => Number(row[getColumnIndex(lowStockHeaders, 'кількість')] || 0) <= 5);
        if (lowStock.length === 0) {
          await interaction.reply({ content: '🟢 Усі товари в наявності.', ephemeral: false });
          return;
        }
        let outputLowStock = '';
        for (let i = 0; i < Math.min(10, lowStock.length); i++) {
          const row = lowStock[i];
          const name = row[getColumnIndex(lowStockHeaders, 'назва')] || '—';
          const quantity = row[getColumnIndex(lowStockHeaders, 'кількість')] || '—';
          outputLowStock += `
• ${name} | Кількість: ${quantity}`;
        }
        const embedLowStock = new EmbedBuilder()
          .setTitle(`⚠️ Мало товару (${lowStock.length})`)
          .setDescription(outputLowStock)
          .setColor(15158332);
        const rowButtons = new ActionRowBuilder()
          .addComponents(
            new ButtonBuilder()
              .setCustomId('download_excel_low_stock')
              .setLabel('Завантажити Excel')
              .setStyle(ButtonStyle.Success)
          );
        await interaction.reply({ embeds: [embedLowStock], components: [rowButtons], ephemeral: false });
        break;

        case 'пошук':
          const field = interaction.options.getString('поле');
          const query = interaction.options.getString('запит').toLowerCase();
        
          const sheetData = await getSheetData();
          const headers = sheetData[0];
          const rows = sheetData.slice(1);
        
          let colIndex = getColumnIndex(headers, field);
          if (colIndex === -1) {
            await interaction.reply({ content: '❌ Невідоме поле для пошуку.', ephemeral: false });
            return;
          }
        
          const isNumericField = ['кількість', 'ціна'].includes(field);
          const results = rows.filter(row => {
            const value = row[colIndex]?.toString().toLowerCase() || '';
            return isNumericField ? Number(value) >= Number(query) : value.includes(query);
          });
        
          if (results.length === 0) {
            return interaction.reply({ content: '🔍 Нічого не знайдено.', ephemeral: false });
          }
          
          cacheSearchResults(interaction.user.id, results);
          
          let currentPage = 0;
          const itemsPerPage = 10;
          
          function generatePageEmbed(results, page) {
            const totalPages = Math.ceil(results.length / itemsPerPage);
            const paginatedResults = results.slice(page * itemsPerPage, (page + 1) * itemsPerPage);
            let output = '| Назва       | Кількість | Ціна |\n|--------------|------------|--------|\n';
          
            for (let i = 0; i < paginatedResults.length && i < itemsPerPage; i++) {
              const row = paginatedResults[i];
              const name = row[6] || '—';
              const quantity = row[3] || '—';
              const price = row[4] || '—';
              output += `| ${name.padEnd(13).slice(0, 13)} | ${quantity} | ${price} |\n`;
            }
          
            return new EmbedBuilder()
              .setTitle(`🔍 Результати пошуку (${results.length})`)
              .setDescription(`\`\`\`md\n${output}\`\`\``)
              .setFooter({ text: `Сторінка ${page + 1}/${totalPages}` })
              .setColor(3066993);
          }
          
          const  = new ActionRowBuilder()
            .addComponents(
              new ButtonBuilder()
                .setCustomId('prev_page')
                .setLabel('⬅️ Попередня')
                .setStyle(ButtonStyle.Secondary),
              new ButtonBuilder()
                .setCustomId('next_page')
                .setLabel('➡️ Наступна')
                .setStyle(ButtonStyle.Secondary),
              new ButtonBuilder()
                .setCustomId('download_excel_search')
                .setLabel('📊 Експортувати')
                .setStyle(ButtonStyle.Success)
            );
          
          await interaction.reply({
            embeds: [generatePageEmbed(results, currentPage)],
            components: [rowButtons],
            ephemeral: false
          });

      case 'пошук-експортовано': {
        const field = interaction.options.getString('поле');
        const query = interaction.options.getString('запит').toLowerCase();
      
        // Завантажуємо таблицю
        const sheetData = await getSheetData();
        const rows = sheetData.slice(1);
        const headers = sheetData[0];
      
        // Визначаємо колонку для пошуку
        let colIndex = getColumnIndex(headers, field);
        if (colIndex === -1) {
          return interaction.reply({ content: '❌ Невідоме поле для пошуку.', ephemeral: false });
        }
      
        // Фільтруємо результати
        const isNumericField = ['кількість', 'ціна'].includes(field);
        const results = rows.filter(row => {
          const value = row[colIndex]?.toString().toLowerCase() || '';
          return isNumericField ? Number(value) >= Number(query) : value.includes(query);
        });
      
        if (results.length === 0) {
          return interaction.reply({ content: '🔍 Нічого не знайдено.', ephemeral: false });
        }
      
        // Формуємо текстовий результат
        let outputSearch = '| Назва       | Кількість | Ціна |\n';
        outputSearch += '|--------------|------------|--------|\n';
      
        for (let i = 0; i < Math.min(10, results.length); i++) {
          const row = results[i];
          const name = row[getColumnIndex(headers, 'назва')] || '—';
          const quantity = row[getColumnIndex(headers, 'кількість')] || '—';
          const price = row[getColumnIndex(headers, 'ціна')] || '—';
          outputSearch += `| ${name.padEnd(13).slice(0, 13)} | ${quantity} | ${price} |\n`;
        }
      
        // Створюємо Embed
        const embedSearch = new EmbedBuilder()
          .setTitle(`🔍 Результати пошуку (${results.length})`)
          .setDescription(`\`\`\`md\n${outputSearch}\`\`\``)
          .setColor(3066993);
      
        // Якщо команда '/пошук' — показуємо кнопку "Завантажити Excel"
        if (interaction.commandName === 'пошук') {
          const rowButtons = new ActionRowBuilder()
            .addComponents(
              new ButtonBuilder()
                .setCustomId('download_excel_search')
                .setLabel('Завантажити Excel')
                .setStyle(ButtonStyle.Primary)
            );
          await interaction.reply({ embeds: [embedSearch], components: [rowButtons], ephemeral: false });
        } 
        // Якщо команда '/пошук-експортовано' — експортуємо автоматично
        else if (interaction.commandName === 'пошук-експортовано') {
          const exportData = [headers, ...results]; // заголовки + результати
          const worksheet = XLSX.utils.aoa_to_sheet(exportData);
          const workbook = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(workbook, worksheet, 'Результати пошуку');
          const filePath = './search_results.xlsx';
          XLSX.writeFile(workbook, filePath);
      
          await interaction.reply({
            content: '📊 Експортуємо результати пошуку:',
            files: [filePath],
            ephemeral: false
          });
        }
        break;
      }

      case 'розумний-пошук': {
        const sheetData = await getSheetData();
        const rows = sheetData.slice(1);
        const headers = sheetData[0];
        const filters = {
          name: interaction.options.getString('номенклатура'),
          client: interaction.options.getString('контрагент'),
          series: interaction.options.getString('серія'),
          priceMin: interaction.options.getNumber('ціна_вище'),
          quantityMin: interaction.options.getNumber('кількість_вище')
        };
      
        const smartResults = rows.filter(row => {
          const nameMatch = !filters.name || row[getColumnIndex(headers, 'назва')]?.toLowerCase().includes(filters.name.toLowerCase());
          const clientMatch = !filters.client || row[getColumnIndex(headers, 'контрагент')]?.toLowerCase().includes(filters.client.toLowerCase());
          const seriesMatch = !filters.series || row[getColumnIndex(headers, 'серія')]?.toLowerCase().includes(filters.series.toLowerCase());
          const priceMatch = !filters.priceMin || Number(row[getColumnIndex(headers, 'ціна')] || 0) >= filters.priceMin;
          const quantityMatch = !filters.quantityMin || Number(row[getColumnIndex(headers, 'кількість')] || 0) >= filters.quantityMin;
      
          return nameMatch && clientMatch && seriesMatch && priceMatch && quantityMatch;
        });
      
        if (smartResults.length === 0) {
          return interaction.reply({ content: '🔍 Нічого не знайдено.', ephemeral: false });
        }
      
        let outputSmartSearch = '| Назва       | Кількість | Ціна |\n';
        outputSmartSearch += '|--------------|------------|--------|\n';
      
        for (let i = 0; i < Math.min(10, smartResults.length); i++) {
          const row = smartResults[i];
          const name = row[getColumnIndex(headers, 'назва')] || '—';
          const quantity = row[getColumnIndex(headers, 'кількість')] || '—';
          const price = row[getColumnIndex(headers, 'ціна')] || '—';
          outputSmartSearch += `| ${name.padEnd(13).slice(0, 13)} | ${quantity} | ${price} |\n`;
        }
      
        const embedSmartSearch = new EmbedBuilder()
          .setTitle(`🔍 Результати розумного пошуку (${smartResults.length})`)
          .setDescription(`\`\`\`md\n${outputSmartSearch}\`\`\``)
          .setColor(3066993);
      
        const rowSmartExport = new ActionRowBuilder()
          .addComponents(
            new ButtonBuilder()
              .setCustomId('download_excel_smart')
              .setLabel('Завантажити Excel')
              .setStyle(ButtonStyle.Success)
          );
      
        // Зберігаємо результати в кеші для експорту
        searchCache[interaction.user.id] = { results: smartResults, headers };
      
        await interaction.reply({ 
          embeds: [embedSmartSearch], 
          components: [rowSmartExport], 
          ephemeral: false 
        });
        break;
      }

      case 'експорт': {
        const sheetData = await getSheetData();
        const exportRows = sheetData || [];
        const worksheet = XLSX.utils.aoa_to_sheet(exportRows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Дані');
        const filePath = './table.xlsx';
        XLSX.writeFile(workbook, filePath);
        await interaction.reply({
          content: '📎 Експортуємо всю таблицю...',
          files: [filePath],
          ephemeral: false
        });
        break;
      }
      case 'help': {
        const helpEmbed = new EmbedBuilder()
          .setTitle('📚 Допомога')
          .setDescription('Ось усі доступні команди:')
          .addFields([
            { name: '/залишки', value: 'Показує загальну кількість і суму товарів', inline: false },
            { name: '/оновити', value: 'Показує останні 10 записів', inline: false },
            { name: '/порожні', value: 'Показує товари, де кількість ≤ 5', inline: false },
            { name: '/пошук [поле] [запит]', value: 'Пошук за полями: назва, серія, контрагент', inline: false },
            { name: '/розумний-пошук', value: 'Пошук за кількома полями', inline: false },
            { name: '/пошук-експортовано [поле] [запит]', value: 'Експортує результати пошуку в Excel', inline: false },
            { name: '!додати [назва] [кількість]', value: 'Додає новий запис через Google Apps Script', inline: false },
            { name: '!експорт', value: 'Експортується вся таблиця', inline: false }
          ])
          .setColor(5763719)
          .setTimestamp();

        const rowHelp = new ActionRowBuilder()
          .addComponents(
            new ButtonBuilder()
              .setLabel('Документація')
              .setURL('https://your-docs-link-here')
              .setStyle(ButtonStyle.Link)
          );

        await interaction.reply({ embeds: [helpEmbed], components: [rowHelp], ephemeral: false });
        break;
      }

      default:
        await interaction.reply({ content: '❌ Невідома команда!', ephemeral: true });
    }
  } catch (err) {
    console.error(err);
    await interaction.reply({ content: '❌ Помилка при завантаженні даних.', ephemeral: true });
  }
});

// Обробка кнопок
client.on('interactionCreate', async interaction => {
  if (!interaction.isButton()) return;

  const userId = interaction.user.id;
  const cachedResults = getCachedResults(userId, 'smart');
  if (!cachedResults && interaction.customId.startsWith('download_excel_smart')) {
    return interaction.reply({ content: '❌ Немає результатів для експорту.', ephemeral: true });
  }

  if (interaction.customId === 'download_excel_smart') {
    try {
      const sheetData = await loadSheetData();
      const headers = sheetData.values[0];
      const exportData = [headers, ...cachedResults];

      const worksheet = XLSX.utils.aoa_to_sheet(exportData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Результати пошуку');

      const filePath = './tmp/smart_search_results.xlsx';
      XLSX.writeFile(workbook, filePath);

      await interaction.update({ content: '📎 Формуємо Excel...', components: [] });
      await interaction.followUp({
        content: '📊 Ось ваша таблиця:',
        files: [filePath],
        ephemeral: false
      });

      // Автоматичне видалення через 10 секунд
      setTimeout(() => {
        require('fs').unlink(filePath, () => {});
      }, 10000);

    } catch (err) {
      console.error('❌ Помилка при експорті:', err);
      await interaction.reply({ content: '⚠️ Не вдалося згенерувати файл.', ephemeral: true });
    }
    return;
  }

  // Обробка сторінок (попередня/наступна)
  if (['prev_page_smart', 'next_page_smart'].includes(interaction.customId)) {
    let currentPage = parseInt(interaction.message.embeds[0].footer.text.match(/(\d+)\/(\d+)/)?.[1] || 1);
    if (interaction.customId === 'prev_page_smart' && currentPage > 1) currentPage--;
    if (interaction.customId === 'next_page_smart' && currentPage * itemsPerPage < cachedResults.length) currentPage++;

    const newEmbed = generatePageEmbed(cachedResults, currentPage, 'smart');
    
    const canPrev = currentPage > 1;
    const canNext = currentPage * itemsPerPage < cachedResults.length;

    const rowButtons = new ActionRowBuilder()
      .addComponents(
        new ButtonBuilder()
          .setCustomId('prev_page_smart')
          .setLabel('⬅️ Попередня')
          .setStyle(ButtonStyle.Secondary)
          .setDisabled(!canPrev),
        new ButtonBuilder()
          .setCustomId('next_page_smart')
          .setLabel('➡️ Наступна')
          .setStyle(ButtonStyle.Secondary)
          .setDisabled(!canNext),
        new ButtonBuilder()
          .setCustomId('download_excel_smart')
          .setLabel('📊 Експортувати')
          .setStyle(ButtonStyle.Success)
      );

    await interaction.update({ embeds: [newEmbed], components: [rowButtons] });
  }
});

setInterval(clearOldFiles, 300000); // кожні 5 хвилин

// Текстові команди
client.on('messageCreate', async msg => {
  if (msg.author.bot) return;

  const args = msg.content.split(' ');

  if (args[0] === '!додати') {
    if (args.length < 3) {
      return msg.reply('Використання: `!додати [назва] [кількість]`');
    }

    const name = args.slice(1, -1).join(' ');
    const quantity = parseInt(args[args.length - 1]);

    if (!name || isNaN(quantity)) {
      return msg.reply('❌ Неправильний формат. Приклад: `!додати ноутбук 5`');
    }

    try {
      const response = await fetch(APP_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, quantity })
      });

      const text = await response.text();
      if (text.trim() === 'OK') {
        msg.reply(`✅ Додано: "${name}" × ${quantity}`);
      } else {
        msg.reply('❌ Не вдалося додати запис.');
      }
    } catch (err) {
      console.error(err);
      msg.reply('❌ Не вдалося відправити запит до Google Apps Script.');
    }
  }

  if (msg.content === '!експорт') {
    try {
      const sheetData = await loadSheetData();
      const worksheet = XLSX.utils.aoa_to_sheet(sheetData.values);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Дані');

      const filePath = './table.xlsx';
      XLSX.writeFile(workbook, filePath);

      await msg.reply({
        content: '📊 Дані експортовано:',
        files: [filePath]
      });
    } catch (err) {
      console.error(err);
      msg.reply('❌ Не вдалося згенерувати файл.');
    }
  }
});

// ────────────────
// 🧠 Кеш для зберігання результатів пошуку
// ────────────────
const searchCache = {};
const CACHE_TTL = 5 * 60 * 1000; // 5 хвилин

function cacheSearchResults(userId, results) {
  searchCache[userId] = {
    results,
    timestamp: Date.now()
  };
}

function getCachedResults(userId) {
  const cached = searchCache[userId];
  if (!cached || Date.now() - cached.timestamp > CACHE_TTL) return null;
  return cached.results;
}

function clearOldFiles() {
  const fs = require('fs');
  const path = require('path');
  const tmpDir = './tmp';
  if (!fs.existsSync(tmpDir)) return;

  fs.readdir(tmpDir, (err, files) => {
    if (err) return console.error('❌ Не вдалося прочитати папку:', err);
    files.forEach(file => {
      const filePath = path.join(tmpDir, file);
      fs.stat(filePath, (err, stats) => {
        if (err) return console.error('❌ Не вдалося отримати статистику файлу:', err);
        if (Date.now() - stats.mtimeMs > CACHE_TTL) {
          fs.unlink(filePath, err => {
            if (err) console.error('❌ Не вдалося видалити файл:', err);
            else console.log(`🗑️ Видалено старий файл: ${file}`);
          });
        }
      });
    });
  });
}

// Логін бота
client.once('ready', async () => {
  console.log(`Бот ${client.user.tag} онлайн!`);
  try {
    await rest.put(Routes.applicationCommands(client.user.id), { body: commands });
    console.log('Slash-команди зареєстровані!');
  } catch (error) {
    console.error('Не вдалося зареєструвати команди:', error);
  }

  // Автоматична перевірка змін кожні 5 хвилин
  setInterval(() => checkForChanges(client), 300000);
});

client.login(process.env.BOT_TOKEN).catch(console.error);