document.addEventListener("DOMContentLoaded", function() {
    const alphabets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    const inputFields = document.getElementById("input-fields");

    alphabets.split("").forEach(letter => {
        let div = document.createElement("div");
        div.innerHTML = `<label>${letter}:</label> <input type="text" id="input-${letter}" placeholder="${letter}">`;
        inputFields.appendChild(div);
    });
});

function saveToExcel() {
    let wb = XLSX.utils.book_new();
    let ws_data = [];
    const alphabets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");

    alphabets.forEach(letter1 => {
        let row = [];
        alphabets.forEach(letter2 => {
            if (letter1 !== letter2) {
                let val1 = document.getElementById(`input-${letter1}`).value || letter1;
                let val2 = document.getElementById(`input-${letter2}`).value || letter2;
                row.push(`${val1} - ${val2}`);
            }
        });
        ws_data.push(row);
    });

    let ws = XLSX.utils.aoa_to_sheet(ws_data);
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

    XLSX.writeFile(wb, "ExcelData.xlsx");
    alert("Excel file saved successfully!");
}

function loadExcelData(event) {
    let file = event.target.files[0];
    if (!file) return;

    let reader = new FileReader();
    reader.onload = function(e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, { type: "array" });
        let sheet = workbook.Sheets[workbook.SheetNames[0]];
        let jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        let table = document.getElementById("excelTable");
        let thead = table.querySelector("thead tr");
        let tbody = table.querySelector("tbody");

        thead.innerHTML = "";
        tbody.innerHTML = "";

        if (jsonData.length > 0) {
            jsonData[0].forEach(header => {
                let th = document.createElement("th");
                th.textContent = header;
                thead.appendChild(th);
            });

            jsonData.slice(1).forEach(row => {
                let tr = document.createElement("tr");
                row.forEach(cell => {
                    let td = document.createElement("td");
                    td.textContent = cell;
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
            });
        }
        alert("Excel data loaded successfully!");
    };

    reader.readAsArrayBuffer(file);
}

function clearFields() {
    document.querySelectorAll("#input-fields input").forEach(input => input.value = "");
}
