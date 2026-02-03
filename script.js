
// Initialize
let transactions = JSON.parse(localStorage.getItem("transactions")) || [];
let settings = JSON.parse(localStorage.getItem("settings")) || {
    companyName: "RAPA CEMENT & GRC",
    companyAddress:
        "Jl. Ngadiretno no. 33, Tamanagung,\nMuntilan, Magelang 56413",
    companyPhone: "08112959125 / 082134567874",
    companyEmail: "rapastone33@gmail.com",
    sjFormat: "XXX",
    poFormat: "PO-YYYY-MM-XXX",
    bankAccount: "BCA A.N NURJAMAL a c :1040257477",
    defaultNoteSJ:
        "Seluruh tegel sudah disertai tambahan\nsebagai cadangan, barang yang sudah dibeli\ntidak bisa dikembalikan lagi",
    defaultNoteNota:
        "Nb. Pembayaran via Bank/Giro/Cek sah bila uang sudah diterima Perusahaan",
    logo: null,
};

// Templates storage state
let excelTemplates = {
    sj: null,
    nota: null,
};

// Set default date to today
document.getElementById("date").valueAsDate = new Date();

// Load settings
function loadSettings() {
    document.getElementById("companyName").value = settings.companyName;
    document.getElementById("companyAddress").value =
        settings.companyAddress;
    document.getElementById("companyPhone").value = settings.companyPhone;
    document.getElementById("companyEmail").value = settings.companyEmail;
    document.getElementById("sjFormat").value = settings.sjFormat;
    document.getElementById("poFormat").value = settings.poFormat;
    document.getElementById("bankAccount").value = settings.bankAccount;
    document.getElementById("defaultNoteSJ").value = settings.defaultNoteSJ;
    document.getElementById("defaultNoteNota").value =
        settings.defaultNoteNota;

    if (settings.logo) {
        document.getElementById("logoPreviewContainer").innerHTML = `
                    <img src="${settings.logo}" class="logo-preview" alt="Logo">
                    <p style="font-size: 12px; color: #28a745;">✓ Logo ter-upload</p>
                `;
    }
}

loadSettings();

// Update button visibility
function updateExcelButtonVisibility() {
    const btn = document.getElementById("btnExcel");
    if (excelTemplates.sj || excelTemplates.nota) {
        btn.style.display = "inline-block";
    } else {
        btn.style.display = "none";
    }
}

// Initialize LocalForage for Excel Templates
async function initTemplateStorage() {
    try {
        excelTemplates.sj = await localforage.getItem("template_sj");
        excelTemplates.nota = await localforage.getItem("template_nota");

        updateTemplateStatus();
    } catch (err) {
        console.error("Error loading templates:", err);
    }
}

function updateTemplateStatus() {
    if (excelTemplates.sj) {
        document.getElementById("statusSJ").innerHTML =
            `<span style="color: #28a745;">✓ Template SJ Tersimpan (${(excelTemplates.sj.byteLength / 1024).toFixed(1)} KB)</span>`;
    } else {
        document.getElementById("statusSJ").innerHTML =
            `<span style="color: #6c757d;">Belum ada template</span>`;
    }

    if (excelTemplates.nota) {
        document.getElementById("statusNota").innerHTML =
            `<span style="color: #28a745;">✓ Template Nota Tersimpan (${(excelTemplates.nota.byteLength / 1024).toFixed(1)} KB)</span>`;
    } else {
        document.getElementById("statusNota").innerHTML =
            `<span style="color: #6c757d;">Belum ada template</span>`;
    }
    updateExcelButtonVisibility();
}

initTemplateStorage();

// Handle Excel Template Upload
async function handleExcelUpload(event, type) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async function (e) {
        const buffer = e.target.result;
        excelTemplates[type] = buffer;

        try {
            await localforage.setItem(`template_${type}`, buffer);
            updateTemplateStatus();
            alert(
                `✅ Template ${type.toUpperCase()} berhasil diupload dan disimpan!`,
            );
        } catch (err) {
            alert("❌ Gagal menyimpan template ke browser.");
            console.error(err);
        }
    };
    reader.readAsArrayBuffer(file);
}

// Handle logo upload
function handleLogoUpload(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function (e) {
            settings.logo = e.target.result;
            localStorage.setItem("settings", JSON.stringify(settings));
            document.getElementById("logoPreviewContainer").innerHTML = `
                        <img src="${e.target.result}" class="logo-preview" alt="Logo">
                        <p style="font-size: 12px; color: #28a745;">✓ Logo ter-upload (klik "Simpan Pengaturan" untuk menyimpan)</p>
                    `;
        };
        reader.readAsDataURL(file);
    }
}

// Generate document number
function generateDocNumber(format) {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, "0");

    // Count transactions this month
    const thisMonthTransactions = transactions.filter((t) => {
        const tDate = new Date(t.date);
        return (
            tDate.getMonth() === now.getMonth() && tDate.getFullYear() === year
        );
    });

    const nextNumber = String(thisMonthTransactions.length + 1).padStart(
        3,
        "0",
    );

    return format
        .replace("YYYY", year)
        .replace("MM", month)
        .replace("XXX", nextNumber);
}

// Update document numbers preview
function updateDocNumbers() {
    document.getElementById("sjNumber").value = generateDocNumber(
        settings.sjFormat,
    );
    document.getElementById("poNumber").value = generateDocNumber(
        settings.poFormat,
    );
}

updateDocNumbers();

// Update item numbers
function updateItemNumbers() {
    const rows = document.getElementById("itemsBody").rows;
    for (let i = 0; i < rows.length; i++) {
        rows[i].querySelector(".item-no").textContent = i + 1;
    }
}

// Tab switching
function switchTab(tab) {
    document
        .querySelectorAll(".tab")
        .forEach((t) => t.classList.remove("active"));
    document
        .querySelectorAll(".tab-content")
        .forEach((c) => c.classList.remove("active"));

    event.target.classList.add("active");
    document.getElementById(tab + "-tab").classList.add("active");

    if (tab === "history") {
        loadHistory();
    }
}

// Add item row
function addItem() {
    const tbody = document.getElementById("itemsBody");
    const row = tbody.insertRow();
    const rowNum = tbody.rows.length;
    row.innerHTML = `
                <td class="item-no">${rowNum}</td>
                <td><input type="text" class="item-code" placeholder="Kode"></td>
                <td><input type="text" class="item-name" placeholder="Nama barang" required></td>
                <td><input type="number" class="item-pcs" value="0" min="0" step="1"></td>
                <td><input type="number" class="item-m2" value="0" min="0" step="0.01"></td>
                <td><input type="number" class="item-price" placeholder="0" min="0" required></td>
                <td><button type="button" class="btn btn-danger btn-small" onclick="removeItem(this)">Hapus</button></td>
            `;
    attachCalculateListeners();
}

// Remove item row
function removeItem(btn) {
    if (document.getElementById("itemsBody").rows.length > 1) {
        btn.closest("tr").remove();
        updateItemNumbers();
        calculateTotal();
    } else {
        alert("Minimal harus ada 1 item");
    }
}

// Calculate total
function calculateTotal() {
    let subtotal = 0;
    const rows = document.getElementById("itemsBody").rows;

    for (let row of rows) {
        const price = parseFloat(row.querySelector(".item-price").value) || 0;
        subtotal += price;
    }

    const dp = parseFloat(document.getElementById("dp").value) || 0;
    const ongkir = parseFloat(document.getElementById("ongkir").value) || 0;
    const kekurangan = subtotal + ongkir - dp;

    document.getElementById("subtotalDisplay").textContent =
        formatCurrency(subtotal);
    document.getElementById("kekuranganDisplay").textContent =
        formatCurrency(kekurangan);
}

// Format currency
function formatCurrency(amount) {
    return (
        "Rp " +
        amount.toLocaleString("id-ID", {
            minimumFractionDigits: 0,
            maximumFractionDigits: 0,
        })
    );
}

// Attach calculate listeners
function attachCalculateListeners() {
    document.querySelectorAll(".item-price").forEach((input) => {
        input.removeEventListener("input", calculateTotal);
        input.addEventListener("input", calculateTotal);
    });

    document
        .getElementById("dp")
        .removeEventListener("input", calculateTotal);
    document.getElementById("dp").addEventListener("input", calculateTotal);

    document
        .getElementById("ongkir")
        .removeEventListener("input", calculateTotal);
    document
        .getElementById("ongkir")
        .addEventListener("input", calculateTotal);
}

attachCalculateListeners();

// Form submission
document
    .getElementById("transactionForm")
    .addEventListener("submit", async function (e) {
        e.preventDefault();

        // Collect data
        const items = [];
        const rows = document.getElementById("itemsBody").rows;

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            items.push({
                no: i + 1,
                code: row.querySelector(".item-code").value,
                name: row.querySelector(".item-name").value,
                pcs: parseFloat(row.querySelector(".item-pcs").value) || 0,
                m2: parseFloat(row.querySelector(".item-m2").value) || 0,
                price: parseFloat(row.querySelector(".item-price").value) || 0,
            });
        }

        const subtotal = items.reduce((sum, item) => sum + item.price, 0);
        const dp = parseFloat(document.getElementById("dp").value) || 0;
        const ongkir =
            parseFloat(document.getElementById("ongkir").value) || 0;
        const kekurangan = subtotal + ongkir - dp;

        const transaction = {
            id: Date.now(),
            date: document.getElementById("date").value,
            sjNumber: generateDocNumber(settings.sjFormat),
            poNumber: generateDocNumber(settings.poFormat),
            customer: {
                name: document.getElementById("customerName").value,
                address: document.getElementById("customerAddress").value,
            },
            items: items,
            subtotal: subtotal,
            dp: dp,
            ongkir: ongkir,
            kekurangan: kekurangan,
            keterangan: document.getElementById("keterangan").value,
        };

        // Save transaction
        transactions.push(transaction);
        localStorage.setItem("transactions", JSON.stringify(transactions));

        // Generate PDFs
        await generateSuratJalanPDF(transaction);
        await generateNotaPDF(transaction);

        alert(
            "✅ Surat Jalan dan Nota Penjualan berhasil dibuat!\n\nFile akan didownload otomatis.",
        );

        // Reset form
        resetForm();
        updateDocNumbers();
    });

// Generate Surat Jalan PDF
async function generateSuratJalanPDF(data) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.width;

    // Header - Company Name (Bold, Large)
    doc.setFontSize(16);
    doc.setFont(undefined, "bold");
    doc.text(settings.companyName, 15, 20);

    // Header - Surat Jalan (Right side)
    doc.setFontSize(14);
    doc.text("SURAT JALAN", pageWidth - 15, 20, { align: "right" });

    // Company Address (smaller)
    doc.setFontSize(9);
    doc.setFont(undefined, "normal");
    const addressLines = settings.companyAddress.split("\n");
    let yPos = 26;
    addressLines.forEach((line) => {
        doc.text(line, 15, yPos);
        yPos += 4;
    });

    // Contact Info
    doc.text(`Telp: ${settings.companyPhone}`, 15, yPos);
    yPos += 4;
    doc.text(`Email: ${settings.companyEmail}`, 15, yPos);

    // NO field (right side)
    doc.setFontSize(10);
    doc.text("NO :", pageWidth - 50, 26);
    doc.rect(pageWidth - 40, 22, 35, 6);

    // Line separator
    doc.setLineWidth(0.5);
    doc.line(15, 45, pageWidth - 15, 45);

    // Document details
    yPos = 52;
    doc.setFontSize(10);
    doc.text(`No Surat: ${data.sjNumber}`, 15, yPos);
    doc.text(`Tanggal: ${formatDate(data.date)}`, pageWidth - 60, yPos);

    yPos += 6;
    doc.text(`No. Po: ${data.poNumber}`, 15, yPos);

    yPos += 8;
    doc.text("Kepada Yth:", 15, yPos);
    yPos += 5;
    doc.setFont(undefined, "bold");
    doc.text(data.customer.name, 15, yPos);
    doc.setFont(undefined, "normal");
    yPos += 5;
    const custAddressLines = doc.splitTextToSize(data.customer.address, 80);
    custAddressLines.forEach((line) => {
        doc.text(line, 15, yPos);
        yPos += 5;
    });

    // Items table
    yPos += 5;
    const tableData = data.items.map((item) => [
        item.no,
        item.code || "-",
        item.name,
        item.pcs > 0 ? item.pcs : "-",
        item.m2 > 0 ? item.m2.toFixed(2) : "-",
        data.keterangan || "",
    ]);

    doc.autoTable({
        startY: yPos,
        head: [["No", "Kode", "Nama Barang", "Qty", "Satuan", "Keterangan"]],
        body: tableData,
        theme: "grid",
        headStyles: {
            fillColor: [30, 60, 114],
            fontSize: 9,
            fontStyle: "bold",
        },
        bodyStyles: {
            fontSize: 9,
        },
        columnStyles: {
            0: { cellWidth: 10, halign: "center" },
            1: { cellWidth: 20 },
            2: { cellWidth: 70 },
            3: { cellWidth: 15, halign: "center" },
            4: { cellWidth: 20, halign: "center" },
            5: { cellWidth: 45 },
        },
    });

    // Note section
    yPos = doc.lastAutoTable.finalY + 10;
    doc.setFontSize(8);
    const noteLines = settings.defaultNoteSJ.split("\n");
    noteLines.forEach((line) => {
        doc.text(line, 15, yPos);
        yPos += 4;
    });

    // Additional note about condition
    yPos += 5;
    doc.setFontSize(9);
    doc.setFont(undefined, "bold");
    doc.text("Barang diterima dengan kondisi baik", 15, yPos);
    doc.setFont(undefined, "normal");

    // Signature section
    yPos += 10;
    const sigY = yPos;

    // Three columns for signatures
    const col1X = 25;
    const col2X = 85;
    const col3X = 145;

    doc.setFontSize(9);
    doc.text("Dibuat Oleh", col1X, sigY, { align: "center" });
    doc.text("Dikirim Oleh", col2X, sigY, { align: "center" });
    doc.text("Diterima Oleh", col3X, sigY, { align: "center" });

    // Signature lines
    doc.line(col1X - 25, sigY + 20, col1X + 25, sigY + 20);
    doc.line(col2X - 25, sigY + 20, col2X + 25, sigY + 20);
    doc.line(col3X - 25, sigY + 20, col3X + 25, sigY + 20);

    doc.setFontSize(8);
    doc.text("Tanggal:", col1X - 25, sigY + 25);
    doc.text("Tanggal:", col2X - 25, sigY + 25);
    doc.text("Tanggal:", col3X - 25, sigY + 25);

    // Add logo if exists
    if (settings.logo) {
        try {
            doc.addImage(settings.logo, "PNG", pageWidth - 50, 10, 35, 14);
        } catch (e) {
            console.log("Logo tidak dapat ditambahkan");
        }
    }

    doc.save(`Surat_Jalan_${data.sjNumber.replace(/\//g, "_")}.pdf`);
}

// Generate Nota Penjualan PDF
async function generateNotaPDF(data) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.width;

    // Header - Company Name (Bold, Large)
    doc.setFontSize(16);
    doc.setFont(undefined, "bold");
    doc.text(settings.companyName, 15, 20);

    // Header - Nota Penjualan (Right side)
    doc.setFontSize(14);
    doc.text("NOTA PENJUALAN", pageWidth - 15, 20, { align: "right" });

    // Company Address (smaller)
    doc.setFontSize(9);
    doc.setFont(undefined, "normal");
    const addressLines = settings.companyAddress.split("\n");
    let yPos = 26;
    addressLines.forEach((line) => {
        doc.text(line, 15, yPos);
        yPos += 4;
    });

    // Contact Info
    doc.text(`Telp: ${settings.companyPhone}`, 15, yPos);
    yPos += 4;
    doc.text(`Email: ${settings.companyEmail}`, 15, yPos);

    // Location info (right side)
    doc.setFontSize(10);
    doc.text(data.customer.name.substring(0, 20), pageWidth - 60, 26);
    doc.text("Magelang", pageWidth - 60, 32);

    // Line separator
    doc.setLineWidth(0.5);
    doc.line(15, 45, pageWidth - 15, 45);

    // Customer details
    yPos = 52;
    doc.setFontSize(10);
    doc.text("Yth:", 15, yPos);
    yPos += 5;
    doc.setFont(undefined, "bold");
    doc.text(data.customer.name, 25, yPos);
    doc.setFont(undefined, "normal");
    yPos += 5;
    doc.text("di:", 15, yPos);
    yPos += 5;
    const custAddressLines = doc.splitTextToSize(data.customer.address, 80);
    custAddressLines.forEach((line) => {
        doc.text(line, 25, yPos);
        yPos += 5;
    });

    yPos += 3;
    doc.text(`NO. PO: ${data.poNumber}`, 15, yPos);

    // Items table
    yPos += 8;
    const tableData = data.items.map((item) => [
        item.no,
        item.name,
        item.pcs > 0 ? item.pcs : "-",
        item.m2 > 0 ? item.m2.toFixed(2) : "-",
        formatCurrency(item.price),
        formatCurrency(item.price),
    ]);

    doc.autoTable({
        startY: yPos,
        head: [["NO", "NAMA BARANG", "PCs", "M²", "HARGA", "JUMLAH"]],
        body: tableData,
        theme: "grid",
        headStyles: {
            fillColor: [30, 60, 114],
            fontSize: 9,
            fontStyle: "bold",
        },
        bodyStyles: {
            fontSize: 9,
        },
        columnStyles: {
            0: { cellWidth: 10, halign: "center" },
            1: { cellWidth: 70 },
            2: { cellWidth: 15, halign: "center" },
            3: { cellWidth: 15, halign: "center" },
            4: { cellWidth: 35, halign: "right" },
            5: { cellWidth: 35, halign: "right" },
        },
    });

    // Summary section
    yPos = doc.lastAutoTable.finalY + 5;
    const summaryX = pageWidth - 75;

    doc.setFontSize(10);
    doc.setFont(undefined, "bold");
    doc.text("Jumlah", summaryX, yPos);
    doc.text(formatCurrency(data.subtotal), pageWidth - 15, yPos, {
        align: "right",
    });

    yPos += 6;
    doc.setFont(undefined, "normal");
    doc.text("DP", summaryX, yPos);
    doc.text(formatCurrency(data.dp), pageWidth - 15, yPos, {
        align: "right",
    });

    yPos += 6;
    doc.text("Ongkir", summaryX, yPos);
    doc.text(formatCurrency(data.ongkir), pageWidth - 15, yPos, {
        align: "right",
    });

    yPos += 6;
    doc.setFont(undefined, "bold");
    doc.text("Kekurangan", summaryX, yPos);
    doc.text(formatCurrency(data.kekurangan), pageWidth - 15, yPos, {
        align: "right",
    });

    // Note section
    yPos += 10;
    doc.setFontSize(8);
    doc.setFont(undefined, "normal");
    doc.text(settings.defaultNoteNota, 15, yPos);
    yPos += 4;
    doc.text(settings.bankAccount, 15, yPos);

    // Signature section
    yPos += 15;
    const sigY = yPos;

    doc.setFontSize(9);
    doc.text("CUSTOMER", 30, sigY, { align: "center" });
    doc.text("RAPA CAST STONE", pageWidth - 40, sigY, { align: "center" });

    // Signature lines
    doc.line(10, sigY + 20, 50, sigY + 20);
    doc.line(pageWidth - 60, sigY + 20, pageWidth - 20, sigY + 20);

    // Add logo if exists
    if (settings.logo) {
        try {
            doc.addImage(settings.logo, "PNG", pageWidth - 50, 10, 35, 14);
        } catch (e) {
            console.log("Logo tidak dapat ditambahkan");
        }
    }

    doc.save(`Nota_${data.sjNumber.replace(/\//g, "_")}.pdf`);
}

// --- EXCEL GENERATION LOGIC ---

async function generateTemplatesExcel() {
    // Collect current form data
    const items = [];
    const rows = document.getElementById("itemsBody").rows;
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        items.push({
            no: i + 1,
            code: row.querySelector(".item-code").value,
            name: row.querySelector(".item-name").value,
            pcs: parseFloat(row.querySelector(".item-pcs").value) || 0,
            m2: parseFloat(row.querySelector(".item-m2").value) || 0,
            price: parseFloat(row.querySelector(".item-price").value) || 0,
        });
    }

    const subtotal = items.reduce((sum, item) => sum + item.price, 0);
    const dp = parseFloat(document.getElementById("dp").value) || 0;
    const ongkir = parseFloat(document.getElementById("ongkir").value) || 0;
    const kekurangan = subtotal + ongkir - dp;

    const data = {
        date: document.getElementById("date").value,
        sjNumber:
            document.getElementById("sjNumber").value ||
            generateDocNumber(settings.sjFormat),
        poNumber:
            document.getElementById("poNumber").value ||
            generateDocNumber(settings.poFormat),
        customer: {
            name: document.getElementById("customerName").value,
            address: document.getElementById("customerAddress").value,
        },
        items: items,
        subtotal: subtotal,
        dp: dp,
        ongkir: ongkir,
        kekurangan: kekurangan,
        keterangan: document.getElementById("keterangan").value,
    };

    const promises = [];
    if (excelTemplates.sj) {
        promises.push(downloadExcel("sj", data));
    }
    if (excelTemplates.nota) {
        promises.push(downloadExcel("nota", data));
    }

    if (promises.length > 0) {
        try {
            await Promise.all(promises);
            alert("✅ Excel berhasil dibuat!");
        } catch (err) {
            console.error(err);
            alert("❌ Gagal membuat Excel. Pastikan template valid.");
        }
    } else {
        alert(
            "Silakan upload template Excel di tab Pengaturan terlebih dahulu.",
        );
    }
}

async function processExcelTemplate(type, data) {
    const templateBuffer = excelTemplates[type];
    if (!templateBuffer) {
        console.log(`No template for ${type}`);
        return null;
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(templateBuffer);

    const mapping = {
        "{{DATE}}": formatDate(data.date),
        "{{TANGGAL}}": formatDate(data.date),
        "{{SJ_NUMBER}}": data.sjNumber,
        "{{NO_SURAT}}": data.sjNumber,
        "{{PO_NUMBER}}": data.poNumber,
        "{{NO_PO}}": data.poNumber,
        "{{CUSTOMER_NAME}}": data.customer.name,
        "{{KEPADA}}": data.customer.name,
        "{{CUSTOMER_ADDRESS}}": data.customer.address,
        "{{ALAMAT}}": data.customer.address,
        "{{SUBTOTAL}}": formatCurrency(data.subtotal),
        "{{DP}}": formatCurrency(data.dp),
        "{{ONGKIR}}": formatCurrency(data.ongkir),
        "{{KEKURANGAN}}": formatCurrency(data.kekurangan),
        "{{NOTES}}": data.keterangan || "",
        "{{KETERANGAN}}": data.keterangan || "",
    };

    workbook.eachSheet((sheet) => {
        // Find and expand item table row
        let itemRowIndex = -1;
        sheet.eachRow((row, rowNum) => {
            row.eachCell((cell) => {
                if (
                    typeof cell.value === "string" &&
                    cell.value.includes("[[NAME]]")
                ) {
                    itemRowIndex = rowNum;
                }
            });
        });

        if (itemRowIndex !== -1 && data.items && data.items.length > 0) {
            const templateRow = sheet.getRow(itemRowIndex);

            // Add rows if more than 1 item
            if (data.items.length > 1) {
                sheet.insertRows(
                    itemRowIndex + 1,
                    Array(data.items.length - 1).fill([]),
                );
            }

            data.items.forEach((item, idx) => {
                const currentRow = sheet.getRow(itemRowIndex + idx);

                // Copy values and styles from template row
                templateRow.eachCell(
                    { includeEmpty: true },
                    (templateCell, colNumber) => {
                        const newCell = currentRow.getCell(colNumber);

                        if (templateCell.style) {
                            newCell.style = JSON.parse(
                                JSON.stringify(templateCell.style),
                            );
                        }

                        if (typeof templateCell.value === "string") {
                            let val = templateCell.value;
                            val = val.replace("[[NO]]", item.no);
                            val = val.replace("[[CODE]]", item.code || "-");
                            val = val.replace("[[KODE]]", item.code || "-");
                            val = val.replace("[[NAME]]", item.name);
                            val = val.replace("[[NAMA]]", item.name);
                            val = val.replace("[[PCS]]", item.pcs || 0);
                            val = val.replace("[[QTY]]", item.pcs || 0);
                            val = val.replace(
                                "[[M2]]",
                                item.m2 ? item.m2.toFixed(2) : "0",
                            );
                            val = val.replace(
                                "[[SATUAN]]",
                                item.m2 ? item.m2.toFixed(2) : "Pcs",
                            );
                            val = val.replace("[[PRICE]]", formatCurrency(item.price));
                            val = val.replace("[[HARGA]]", formatCurrency(item.price));
                            newCell.value = val;
                        }
                    },
                );
            });
        }

        // Replace simple tags in all cells
        sheet.eachRow((row) => {
            row.eachCell((cell) => {
                if (typeof cell.value === "string") {
                    let newVal = cell.value;
                    for (const [tag, value] of Object.entries(mapping)) {
                        if (newVal.includes(tag)) {
                            newVal = newVal.split(tag).join(value);
                        }
                    }
                    cell.value = newVal;
                }
            });
        });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    return buffer;
}

async function downloadExcel(type, data) {
    const buffer = await processExcelTemplate(type, data);
    if (!buffer) return;

    const filename =
        type === "sj"
            ? `Surat_Jalan_${data.sjNumber.replace(/\//g, "_")}.xlsx`
            : `Nota_${data.sjNumber.replace(/\//g, "_")}.xlsx`;

    const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
    URL.revokeObjectURL(link.href);
}

// Format date
function formatDate(dateString) {
    const date = new Date(dateString);
    const months = [
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "Mei",
        "Jun",
        "Jul",
        "Agu",
        "Sep",
        "Okt",
        "Nov",
        "Des",
    ];
    return `${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear()}`;
}

// Reset form
function resetForm() {
    document.getElementById("transactionForm").reset();
    document.getElementById("date").valueAsDate = new Date();

    // Reset items table to 1 row
    const tbody = document.getElementById("itemsBody");
    tbody.innerHTML = `
                <tr>
                    <td class="item-no">1</td>
                    <td><input type="text" class="item-code" placeholder="Kode"></td>
                    <td><input type="text" class="item-name" placeholder="Nama barang" required></td>
                    <td><input type="number" class="item-pcs" value="0" min="0" step="1"></td>
                    <td><input type="number" class="item-m2" value="0" min="0" step="0.01"></td>
                    <td><input type="number" class="item-price" placeholder="0" min="0" required></td>
                    <td><button type="button" class="btn btn-danger btn-small" onclick="removeItem(this)">Hapus</button></td>
                </tr>
            `;

    attachCalculateListeners();
    calculateTotal();
}

// Load history
function loadHistory() {
    const content = document.getElementById("historyContent");

    if (transactions.length === 0) {
        content.innerHTML = `
                    <div class="empty-state">
                        <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                        </svg>
                        <h3>Belum Ada Transaksi</h3>
                        <p>Buat transaksi pertama Anda di tab "Buat Transaksi"</p>
                    </div>
                `;
        return;
    }

    let html = `
                <table class="history-table">
                    <thead>
                        <tr>
                            <th>Tanggal</th>
                            <th>No. Surat Jalan</th>
                            <th>No. PO</th>
                            <th>Customer</th>
                            <th>Total</th>
                            <th>Status</th>
                            <th>Aksi</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

    transactions
        .slice()
        .reverse()
        .forEach((t) => {
            html += `
                    <tr>
                        <td>${formatDate(t.date)}</td>
                        <td>${t.sjNumber}</td>
                        <td>${t.poNumber}</td>
                        <td>${t.customer.name}</td>
                        <td>${formatCurrency(t.kekurangan)}</td>
                        <td><span class="status-badge status-completed">Selesai</span></td>
                        <td>
                            <button class="btn btn-primary btn-small" onclick="regeneratePDFs(${t.id})">Download Ulang</button>
                        </td>
                    </tr>
                `;
        });

    html += `
                    </tbody>
                </table>
            `;

    content.innerHTML = html;
}

// Regenerate PDFs
async function regeneratePDFs(id) {
    const transaction = transactions.find((t) => t.id === id);
    if (transaction) {
        await generateSuratJalanPDF(transaction);
        await generateNotaPDF(transaction);
        alert("✅ Surat Jalan dan Nota berhasil didownload ulang!");
    }
}

// Save settings
function saveSettings() {
    settings = {
        companyName: document.getElementById("companyName").value,
        companyAddress: document.getElementById("companyAddress").value,
        companyPhone: document.getElementById("companyPhone").value,
        companyEmail: document.getElementById("companyEmail").value,
        sjFormat: document.getElementById("sjFormat").value,
        poFormat: document.getElementById("poFormat").value,
        bankAccount: document.getElementById("bankAccount").value,
        defaultNoteSJ: document.getElementById("defaultNoteSJ").value,
        defaultNoteNota: document.getElementById("defaultNoteNota").value,
        logo: settings.logo, // Keep existing logo
    };

    localStorage.setItem("settings", JSON.stringify(settings));
    updateDocNumbers();
    alert("✅ Pengaturan berhasil disimpan!");
}

// Clear all data
function clearAllData() {
    if (
        confirm(
            "⚠️ PERINGATAN!\n\nApakah Anda yakin ingin menghapus SEMUA data?\nData yang dihapus tidak dapat dikembalikan!",
        )
    ) {
        if (
            confirm(
                "Konfirmasi sekali lagi. Yakin ingin menghapus semua data transaksi?",
            )
        ) {
            localStorage.clear();
            transactions = [];
            location.reload();
        }
    }
}
