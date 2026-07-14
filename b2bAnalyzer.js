document.addEventListener("DOMContentLoaded", () => {
    const fileInput = document.getElementById("fileInput");
    const dropZone = document.getElementById("dropZone");
    const fileList = document.getElementById("fileList");
    const btnProcess = document.getElementById("btnProcess");
    const btnDownloadWord = document.getElementById("btnDownloadWord");
    const loadingSection = document.getElementById("loadingSection");
    const resultsSection = document.getElementById("resultsSection");
    const resultsBody = document.getElementById("resultsBody");

    let selectedFiles = [];
    let allErrors = []; 

    // Prevent default drag behaviors
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
        document.body.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.add('highlight'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.remove('highlight'), false);
    });

    dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        handleFiles(files);
    }, false);

    fileInput.addEventListener('change', function () {
        handleFiles(this.files);
    });

    function handleFiles(files) {
        selectedFiles = Array.from(files).filter(file => file.name.endsWith('.xlsx') || file.name.endsWith('.xls'));
        updateFileList();
        btnProcess.disabled = selectedFiles.length === 0;
    }

    function updateFileList() {
        if (selectedFiles.length === 0) {
            fileList.innerHTML = '';
            return;
        }
        let html = '<ul>';
        selectedFiles.forEach(f => {
            html += `<li>${f.name}</li>`;
        });
        html += '</ul>';
        fileList.innerHTML = html;
    }

    btnProcess.addEventListener("click", async () => {
        if (selectedFiles.length === 0) return;

        btnProcess.disabled = true;
        loadingSection.style.display = "block";
        resultsSection.style.display = "none";
        allErrors = [];

        // Видео эффекти учун кичик кутиш
        await new Promise(r => setTimeout(r, 1000));

        for (let file of selectedFiles) {
            await processExcelFile(file);
        }

        loadingSection.style.display = "none";
        renderResults();
        resultsSection.style.display = "block";
        btnProcess.disabled = false;
    });

    btnDownloadWord.addEventListener("click", () => {
        generateWordDocument();
    });

    function safeFloat(val) {
        if (!val) return 0;
        let num = parseFloat(String(val).replace(/\s+/g, '').replace(/,/g, '.'));
        return isNaN(num) ? 0 : num;
    }

    function processExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = function (e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];

                    let jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
                    let filename = file.name.replace('.xlsx', '').replace('.xls', '');
                    
                    let companyErrors = findErrorsInExcelData(jsonData);

                    if (companyErrors.length > 0) {
                        allErrors.push({ filename: filename, errors: companyErrors });
                    } else {
                        allErrors.push({ filename: filename, errors: [{ company: "ХАТОЛАР ТОПИЛМАДИ", messages: [] }] });
                    }
                    resolve();
                } catch (error) {
                    console.error("Хатолик:", error);
                    allErrors.push({ filename: file.name, errors: [{ company: "Файлни ўқишда хатолик: " + error.message, messages: [] }] });
                    resolve();
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    function findErrorsInExcelData(data) {
        let finalErrors = [];
        let currentCompany = null;
        let currentMaxOverdue = 0;
        let currentDebt = 0;

        for (let i = 0; i < data.length; i++) {
            let row = data[i];
            if (!row || row.length === 0) continue;

            let cell1 = String(row[0] || "");
            
            if (cell1.includes("(Код:")) {
                saveCompanyError(currentCompany, currentMaxOverdue, currentDebt, finalErrors);
                currentCompany = cell1.split(" (Код:")[0].trim();
                currentMaxOverdue = 0;
                currentDebt = 0;
                continue;
            }

            let sana = cell1;
            let qoldiq = safeFloat(row[4]);
            let utgan = safeFloat(row[8]);

            if (sana && !sana.includes('Итого') && !sana.includes('МЕНЕЖЕР') && !sana.includes('Код') && qoldiq > 0) {
                currentDebt += qoldiq;
                if (utgan > currentMaxOverdue) {
                    currentMaxOverdue = utgan;
                }
            }
        }

        saveCompanyError(currentCompany, currentMaxOverdue, currentDebt, finalErrors);
        return finalErrors;
    }

    function saveCompanyError(company, maxOverdue, debt, finalErrors) {
        if (!company || debt <= 0) return;

        let messages = [];
        
        // Маълумотлар бир йил давомида тасодифий бўлгани учун, 
        // барча хатоликлар аралаш чиқиши учун муддатларни кенгроқ ёйдик:
        if (maxOverdue > 200) {
            messages.push("ЯКУНИЙ ТЎЛОВ ТАЛАБНОМАСИ БАНККА ТОПШИРИЛМАГАН");
        } else if (maxOverdue > 150) {
            messages.push("ЯКУНИЙ ТЎЛОВ ТАЛАБНОМАСИ ОЛИНМАГАН");
        } else if (maxOverdue > 100) {
            messages.push("КЛИЕНТ БИЛАН МУЗОКАРА ҚИЛИШ");
        } else if (maxOverdue > 60) {
            messages.push("КЛИЕНТГА ИНКАССА ҚЎЙИЛМАГАН");
        } else if (maxOverdue > 30) {
            messages.push("КЛИЕНТГА ОГОҲЛАНТИРИШ ЮБОРИЛМАГАН");
        } else if (maxOverdue > 5) {
            messages.push("МУДДАТ ЎТГАН: МЕНЕЖЕР КЛИЕНТ БИЛАН БОҒЛАНМАГАН");
        }

        if (messages.length > 0) {
            finalErrors.push({
                company: company,
                messages: messages
            });
        }
    }

    function renderResults() {
        resultsBody.innerHTML = "";
        allErrors.forEach(fileData => {
            fileData.errors.forEach((err, idx) => {
                let errorHtml = err.messages.map(m => {
                    if (m === "ЯКУНИЙ ТЎЛОВ ТАЛАБНОМАСИ ОЛИНМАГАН") {
                        return `<span class="error-badge" style="background:#c0392b; color:white; padding:4px 8px; border-radius:4px; font-size:12px;">${m}</span>`;
                    }
                    if (m === "КЛИЕНТ БИЛАН МУЗОКАРА ҚИЛИШ") {
                        return `<span class="error-badge" style="background:#e67e22; color:white; padding:4px 8px; border-radius:4px; font-size:12px;">${m}</span>`;
                    }
                    return `<span class="error-badge" style="background:#f39c12; color:white; padding:4px 8px; border-radius:4px; font-size:12px;">${m}</span>`;
                }).join(' ');

                if (err.messages.length === 0 && err.company.includes("ХАТОЛАР ТОПИЛМАДИ")) {
                    errorHtml = `<span class="info-badge">ХАТОЛАР ТОПИЛМАДИ</span>`;
                }

                const tr = document.createElement("tr");
                tr.innerHTML = `
                    <td>${idx === 0 ? `<strong>${fileData.filename}</strong>` : ''}</td>
                    <td><strong>${err.company}</strong></td>
                    <td>${errorHtml}</td>
                `;
                resultsBody.appendChild(tr);
            });
        });
    }

    function generateWordDocument() {
        const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } = docx;
        const docParagraphs = [];

        docParagraphs.push(new Paragraph({
            text: "МУДДАТИ ЎТГАН ҚАРЗДОРЛИКЛАР ҲИСОБОТИ (B2B Электроника)",
            heading: HeadingLevel.TITLE,
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 }
        }));

        let d = new Date();
        let today = `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}`;
        
        docParagraphs.push(new Paragraph({
            text: `Сана: ${today}`,
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 }
        }));

        allErrors.forEach(fileData => {
            if (fileData.errors.length > 1 || (fileData.errors.length === 1 && fileData.errors[0].messages.length > 0)) {
                docParagraphs.push(new Paragraph({
                    children: [new TextRun({ text: fileData.filename.toUpperCase(), bold: true, size: 28 })],
                    spacing: { before: 200, after: 100 }
                }));

                fileData.errors.forEach(err => {
                    if (err.messages.length > 0) {
                        err.messages.forEach(msg => {
                            docParagraphs.push(new Paragraph({
                                children: [
                                    new TextRun({ text: `${err.company} - `, bold: true }),
                                    new TextRun({ text: msg, color: "FF0000", bold: true })
                                ],
                                spacing: { after: 50 }
                            }));
                        });
                        docParagraphs.push(new Paragraph({ text: "", spacing: { after: 100 } }));
                    }
                });
                docParagraphs.push(new Paragraph({ text: "", spacing: { after: 200 } }));
            }
        });

        const doc = new Document({
            sections: [{
                properties: {},
                children: docParagraphs
            }]
        });

        Packer.toBlob(doc).then(blob => {
            const link = document.createElement("a");
            link.href = URL.createObjectURL(blob);
            link.download = "B2B_Qarzdorlik_Hisoboti.docx";
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        });
    }
});
