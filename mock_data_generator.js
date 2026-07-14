const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// Папкани яратиш
const DATA_DIR = path.join(__dirname, 'MockData');
if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR);
}

// 20 та филиал (Тошкент ва вилоятлар) кириллчада
const REGIONS = [
    "Тошкент_Чилонзор", "Тошкент_Юнусобод", "Тошкент_Сергели", "Тошкент_Мирзо_Улуғбек", "Тошкент_Яккасарой",
    "Тошкент_Миробод", "Тошкент_Олмазор", "Тошкент_Учтепа", "Тошкент_Шайхонтоҳур", "Тошкент_Бектемир",
    "Самарқанд_Марказ", "Фарғона_Марказ", "Андижон_Марказ", "Наманган_Марказ", "Бухоро_Марказ",
    "Хоразм_Урганч", "Навоий_Марказ", "Қашқадарё_Қарши", "Сурхондарё_Термиз", "Жиззах_Марказ"
];

// Технологик фирма номларини генерация қилиш учун сўзлар
const PREFIXES = ["Смарт", "Рақамли", "Техно", "Мега", "Глобал", "Инновацион", "Эл", "АйТи", "Мобил", "Алоқа", "Супер", "Оптом", "Етакчи", "Замонавий", "Тезкор"];
const BASES = ["Электроникс", "Тех", "Дунёси", "Олами", "Трейд", "Бизнес", "Маркет", "База", "Гаджет", "Девайс", "Савдо", "Опт", "Дистрибьюшн", "Системс"];
const TYPES = ["МЧЖ", "ХК", "ОК", "ҚК"];

// Биз келишган янги устунлар номлари
const COLUMNS = [
    { header: 'Ҳужжат санаси', key: 'sana', width: 15 },
    { header: 'Ҳужжат суммаси', key: 'summa', width: 20 },
    { header: 'Олдиндан тўлов %', key: 'foiz', width: 18 },
    { header: 'Тўланган сумма', key: 'tulangan', width: 20 },
    { header: 'Қарз қолдиғи', key: 'qoldiq', width: 20 },
    { header: 'Рухсат этилган муддат (кун)', key: 'muddat_kun', width: 25 },
    { header: 'Тўлов муддати', key: 'tulov_muddati', width: 15 },
    { header: 'Фарқ (кун)', key: 'farq', width: 15 },
    { header: 'Муддати ўтган (кун)', key: 'utgan', width: 25 },
    { header: 'Огоҳлантириш хати', key: 'ogoh', width: 20 },
    { header: 'Судга юбориш санаси', key: 'sud_yuborish', width: 22 },
    { header: 'Суд мажлиси куни', key: 'sud_kuni', width: 20 }
];

function getRandomInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}

function generateCompanyName() {
    const prefix = PREFIXES[getRandomInt(0, PREFIXES.length - 1)];
    const base = BASES[getRandomInt(0, BASES.length - 1)];
    const type = TYPES[getRandomInt(0, TYPES.length - 1)];
    return `${prefix}-${base} ${type}`;
}

function generateDate(start, end) {
    return new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
}

function formatDate(date) {
    let d = date.getDate().toString().padStart(2, '0');
    let m = (date.getMonth() + 1).toString().padStart(2, '0');
    let y = date.getFullYear();
    return `${d}.${m}.${y}`;
}

function addDays(date, days) {
    let result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
}

async function generateExcelFile(regionIndex) {
    const regionName = REGIONS[regionIndex];
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Қарздорлик');

    // Тепа қисмидаги сарлавҳа
    sheet.addRow([]);
    sheet.addRow([`Электроника савдоси бўйича қарздорлик қолдиғи: ${regionName.toUpperCase().replace('_', ' ')} (01.01.2025 - 31.12.2025)`]);
    sheet.addRow([]);

    // 4-қатор: Устунлар номлари
    const headerRow = sheet.addRow(COLUMNS.map(c => c.header));
    
    // Шапкани оч мовий (Pale Blue) корпоратив рангга бўяш
    headerRow.eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE3F2FD' } // Оч корпоратив мовий ранг
        };
        cell.font = { bold: true };
        cell.border = {
            top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
        };
    });

    // Устунлар кенглигини ўрнатиш
    COLUMNS.forEach((col, idx) => {
        sheet.getColumn(idx + 1).width = col.width;
    });

    const currentDate = new Date(2025, 11, 31); // Ҳисобот санаси (31 Дек 2025)

    let rowIndex = 5;
    
    // Менежер гуруҳи
    sheet.addRow([`МЕНЕЖЕР: АБДУЛЛОҲ ТЕШАБОЕВ (${regionName.replace('_', ' ')})`]);
    sheet.getRow(rowIndex).font = { bold: true };
    rowIndex++;

    let totalRowsTarget = 1000;
    let currentRowCount = 0;

    // Мижозларни камида 1000 қатор бўлгунча генерация қилиш
    while (currentRowCount < totalRowsTarget) {
        let clientName = generateCompanyName();
        let clientCode = getRandomInt(10000, 99999);
        let inn = getRandomInt(300000000, 399999999);
        
        // Мижознинг бош қатори
        sheet.addRow([`${clientName} (Код: ${clientCode}) (D) (ИНН: ${inn})`]);
        sheet.getRow(rowIndex).font = { bold: true };
        rowIndex++;

        let numInvoices = getRandomInt(1, 5); // Битта мижознинг 1 дан 5 гача накладнойи бор
        let clientTotalSum = 0;
        let clientTotalPaid = 0;
        let clientTotalDebt = 0;

        for (let i = 0; i < numInvoices; i++) {
            let docDate = generateDate(new Date(2025, 0, 1), new Date(2025, 11, 1));
            let summa = getRandomInt(10, 500) * 100000; // 1 млн дан 50 млн гача электроника суммаси
            let foiz = 15; // 15% келишилган предоплата
            let tulangan = summa * 0.15; 
            
            // Баъзилар кўпроқ тўлаган бўлиши мумкин
            if (Math.random() > 0.6) {
                tulangan += getRandomInt(1, Math.floor(summa * 0.5 / 100000)) * 100000;
            }
            
            let qoldiq = summa - tulangan;
            let muddat_kun = 35; // 35 кун муддат
            let tulov_muddati = addDays(docDate, muddat_kun);
            
            let farq = Math.floor((tulov_muddati - currentDate) / (1000 * 60 * 60 * 24));
            let utgan = farq < 0 ? Math.abs(farq) : '';

            clientTotalSum += summa;
            clientTotalPaid += tulangan;
            clientTotalDebt += qoldiq;

            let row = sheet.addRow([
                formatDate(docDate),
                summa,
                foiz,
                tulangan,
                qoldiq,
                muddat_kun,
                formatDate(tulov_muddati),
                farq,
                utgan,
                '', // Огоҳлантириш хати
                '', // Судга юбориш
                ''  // Суд куни
            ]);
            
            // Сонларни форматлаш (мингталик ажратувчи билан)
            row.getCell(2).numFmt = '#,##0';
            row.getCell(4).numFmt = '#,##0';
            row.getCell(5).numFmt = '#,##0';

            rowIndex++;
            currentRowCount++;
        }
        
        // Мижоз бўйича "Итого" (Жами)
        let itogoRow = sheet.addRow(['Итого', clientTotalSum, '', clientTotalPaid, clientTotalDebt]);
        itogoRow.font = { bold: true };
        itogoRow.getCell(2).numFmt = '#,##0';
        itogoRow.getCell(4).numFmt = '#,##0';
        itogoRow.getCell(5).numFmt = '#,##0';
        rowIndex++;
    }

    const filePath = path.join(DATA_DIR, `${regionName}.xlsx`);
    await workbook.xlsx.writeFile(filePath);
    console.log(`[МУВАФФАҚИЯТ] ${regionName}.xlsx яратилди (${currentRowCount} қатор накладной).`);
}

async function main() {
    console.log("==========================================");
    console.log("БАЗАНИ ГЕНЕРАЦИЯ ҚИЛИШ БОШЛАНДИ...");
    console.log("==========================================");
    for (let i = 0; i < REGIONS.length; i++) {
        await generateExcelFile(i);
    }
    console.log("==========================================");
    console.log("БАРЧА 20 ТА ФАЙЛ МУВАФФАҚИЯТЛИ ЯРАТИЛДИ!");
    console.log("==========================================");
}

main().catch(err => console.error(err));
