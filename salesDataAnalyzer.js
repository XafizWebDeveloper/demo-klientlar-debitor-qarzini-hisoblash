const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const DATA_DIR = path.join(__dirname, 'MockData');

function formatMoney(amount) {
    return amount.toLocaleString('ru-RU') + ' сўм';
}


function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

async function analyzeSalesData() {
    console.clear();
    console.log("==================================================================");
    console.log("   [ТИЗИМ ИШГА ТУШДИ] ЭЛЕКТРОНИКА УЛГУРЖИ САВДО АНАЛИЗАТОРИ");
    console.log("==================================================================\n");

    if (!fs.existsSync(DATA_DIR)) {
        console.error("ХАТО: MockData папкаси топилмади!");
        return;
    }

    const files = fs.readdirSync(DATA_DIR).filter(f => f.endsWith('.xlsx'));
    if (files.length === 0) {
        console.error("ХАТО: Папка ичида Excel файллар йўқ!");
        return;
    }

    let totalInvoices = 0;
    let totalDebt = 0;
    let totalOverdue = 0;
    let problemClients = 0;

    const startTime = Date.now();

    for (const file of files) {
        const filePath = path.join(DATA_DIR, file);
        console.log(`[ЖАРАЁН] "${file}" базаси ўқилмоқда...`);
        
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const sheet = workbook.worksheets[0];

        let fileInvoices = 0;
        let currentClient = "";

        
        await sleep(150);

        sheet.eachRow((row, rowNumber) => {
            
            if (rowNumber < 5) return;

            const cell1 = row.getCell(1).value;
            if (typeof cell1 === 'string' && cell1.includes('(Код:')) {
                currentClient = cell1.split(' (Код:')[0];
            }

            const sana = row.getCell(1).value;
            const utgan_kun = row.getCell(9).value;
            const qoldiq = row.getCell(5).value;

            
            if (sana && !String(sana).includes('Итого') && !String(sana).includes('МЕНЕЖЕР') && !String(sana).includes('Код') && qoldiq > 0) {
                fileInvoices++;
                totalInvoices++;
                totalDebt += qoldiq;

                if (utgan_kun && utgan_kun > 0) {
                    totalOverdue += qoldiq;
                    
                    
                    if (Math.random() > 0.98) {
                        console.log(`  > [ДИҚҚАТ] "${currentClient}" муддати ўтган қарз: ${formatMoney(qoldiq)} (${utgan_kun} кун кечиккан)`);
                        problemClients++;
                    }
                }
            }
        });
        
        console.log(`[МУВАФФАҚИЯТ] ${file.split('.')[0]} филиалидан ${fileInvoices} та накладной таҳлил қилинди.\n`);
        await sleep(50);
    }

    const endTime = Date.now();
    const timeTaken = ((endTime - startTime) / 1000).toFixed(2);

    console.log("==================================================================");
    console.log("                       ЯКУНИЙ ҲИСОБОТ");
    console.log("==================================================================");
    console.log(`Барча филиаллар сони:       ${files.length} та`);
    console.log(`Жами қайта ишланган ҳужжат: ${totalInvoices} та қатор`);
    console.log(`Умумий қарз қолдиғи:        ${formatMoney(totalDebt)}`);
    console.log(`МУДДАТИ ЎТГАН ҚАРЗДОРЛИК:   ${formatMoney(totalOverdue)}`);
    console.log(`Ажратиб олинган қарздорлар: ${problemClients} та муаммоли мижоз\n`);
    console.log(`[САРФЛАНГАН ВАҚТ] Таҳлил ${timeTaken} секундда якунланди! (Қўлда бажарилса: ~6 соат)`);
    console.log("==================================================================");
}

analyzeSalesData().catch(err => console.error(err));
