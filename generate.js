import ExcelJS from 'exceljs';
import fetch from 'node-fetch';
import fs from 'fs';

const data = JSON.parse(fs.readFileSync('data.json', 'utf8'));

const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet('Наряды');

ws.addRow([
  'Транспортные сутки',
  '№ маршрута',
  'Идентификатор выхода ТС',
  'Гаражный номер ТС',
  'ГРЗ',
  'Таб. № водителя',
  'Время начала',
  'Время окончания'
]);

let i = 1;
for (const r of data.rows) {
  ws.addRow([
    data.date,
    '120к',
    `120к_1_${i++}`,
    r.garage,
    r.grz,
    '',
    '04:00',
    '02:00'
  ]);
}

await wb.xlsx.writeFile('Наряды.xlsx');

/* ОТПРАВКА В TELEGRAM */
const form = new FormData();
form.append('chat_id', data.chatId);
form.append(
  'document',
  fs.createReadStream('Наряды.xlsx')
);

await fetch(
  `https://api.telegram.org/bot${data.token}/sendDocument`,
  { method: 'POST', body: form }
);
