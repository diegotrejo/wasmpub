// wwwroot/js/excelInterop.js
function autosizeColumnsFromData(rows) {
    if (!rows || rows.length === 0) return [];
    const headers = Object.keys(rows[0]);
    const cols = headers.map(h => ({ wch: Math.max(10, (h ?? '').toString().length) }));
    for (const row of rows) {
        headers.forEach((h, i) => {
            const v = row[h];
            const len = v == null ? 0 : (v instanceof Date ? 19 : v.toString().length);
            cols[i].wch = Math.max(cols[i].wch, len + 2);
        });
    }
    return cols;
}

function colLetter(n) {
    let s = "";
    while (n >= 0) {
        s = String.fromCharCode((n % 26) + 65) + s;
        n = Math.floor(n / 26) - 1;
    }
    return s;
}

function detectAndFormat(ws, headerMap, range) {
    const intFmt = "#,##0";
    const decFmt = "#,##0.00";
    const dateFmt = "yyyy-mm-dd";
    const dateTimeFmt = "yyyy-mm-dd hh:mm:ss";

    for (const colName of Object.keys(headerMap)) {
        const ci = headerMap[colName];
        const colL = colLetter(ci);

        for (let r = range.s.r + 1; r <= range.e.r; r++) {
            const addr = colL + (r + 1);
            const cell = ws[addr];
            if (!cell) continue;
            let v = cell.v;

            // Detecta nulo
            if (v == null) continue;

            // Detecta números
            if (typeof v === "number") {
                if (Number.isInteger(v)) {
                    cell.t = "n";
                    cell.v = v;
                    cell.z = intFmt;
                } else {
                    cell.t = "n";
                    cell.v = v;
                    cell.z = decFmt;
                }
                continue;
            }

            // Detecta fechas ISO (de Blazor vienen así) o Date
            let d = null;
            if (v instanceof Date) {
                d = v;
            } else if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(v)) {
                d = new Date(v);
            }
            if (d && !isNaN(d.getTime())) {
                cell.t = "d";
                cell.v = d;
                if (d.getHours() === 0 && d.getMinutes() === 0 && d.getSeconds() === 0) {
                    cell.z = dateFmt;
                } else {
                    cell.z = dateTimeFmt;
                }
                continue;
            }

            // Resto: texto
            cell.t = "s";
            cell.v = String(v);
        }
    }
}

window.ExportExcel = {
    fromObjects: function (fileName, sheetName, rows) {
        if (!rows || rows.length === 0) {
            console.warn("No hay datos");
            return;
        }

        const ws = XLSX.utils.json_to_sheet(rows, { cellDates: true, raw: true });
        const range = XLSX.utils.decode_range(ws["!ref"]);
        ws["!autofilter"] = { ref: XLSX.utils.encode_range(range) };
        ws["!cols"] = autosizeColumnsFromData(rows);

        const headers = Object.keys(rows[0] || {});
        const headerMap = {};
        headers.forEach((h, i) => headerMap[h] = i);

        detectAndFormat(ws, headerMap, range);

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, sheetName || "Datos");

        const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array", cellDates: true });
        const blob = new Blob([wbout], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        });

        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = fileName || "datos.xlsx";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(a.href), 0);
    }
};
