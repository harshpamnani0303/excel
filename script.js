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
    const alphabets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split(""); // A-Z tak list

    let name_mapping = {}; // ✅ Custom names mapping

    // ✅ User Inputs Se Names Lena
    alphabets.forEach(letter => {
        let inputElement = document.getElementById(`input-${letter}`);
        if (inputElement) {
            name_mapping[letter] = inputElement.value.trim() || letter; // Agar blank ho to default letter use karein
        } else {
            name_mapping[letter] = letter;
        }
    });

    let columnWidths = []; // ✅ Column width track karne ke liye

    // ✅ Column-wise data arrange karna
    for (let col_index = 0; col_index < alphabets.length; col_index++) { 
        let column = []; // Har column ke liye ek list
        let char1 = alphabets[col_index]; // Fixed column character
        let maxCellLength = 0; // ✅ Track max width for column

        for (let row_index = 0; row_index < alphabets.length; row_index++) { 
            let char2 = alphabets[row_index]; // Changing row character

            // ✅ Self-Pairing Avoid (A - A, B - B)
            if (char1 !== char2) {
                let name1 = name_mapping[char1];
                let name2 = name_mapping[char2];

                let cell_value = `${name1} - ${name2}`;
                column.push(cell_value);

                // ✅ Maximum cell length track karein for width adjustment
                maxCellLength = Math.max(maxCellLength, cell_value.length);
            }
        }

        ws_data.push(column); // ✅ Column-wise data store karna
        columnWidths.push({ wch: maxCellLength + 2 }); // ✅ Width set (extra padding ke saath)
    }

    // ✅ Transpose Data to Get Correct Column Format
    let maxRows = Math.max(...ws_data.map(col => col.length));
    let ws_data_transposed = Array.from({ length: maxRows }, (_, row) =>
        ws_data.map(col => col[row] || "") // Fill missing cells with empty strings
    );

    let ws = XLSX.utils.aoa_to_sheet(ws_data_transposed);

    // ✅ Auto Adjust Columns Width
    ws["!cols"] = columnWidths;

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
