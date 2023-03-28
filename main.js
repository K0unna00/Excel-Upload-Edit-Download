let tableBody = document.querySelector(".table-body");
let tableHead = document.querySelector(".table-head");
let downloadBtn = document.querySelector(".downloadBtn");
let headArr = [];
let finalData = '';
function upload() {
    let files = document.getElementById('file_upload').files;
    if (files.length == 0) {
        alert("Please choose any file...");
        return;
    }
    let filename = files[0].name;
    let extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();
    if (extension == '.XLS' || extension == '.XLSX') {
        excelFileToJSON(files[0]);
    } else {
        alert("Please select a valid excel file.");
    }
}

function excelFileToJSON(file) {
    try {
        let reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function (e) {

            let data = e.target.result;
            let workbook = XLSX.read(data, {
                type: 'binary'
            });

            let result = {};
            workbook.SheetNames.forEach(function (sheetName) {
                let roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
                if (roa.length > 0) {
                    result[sheetName] = roa;
                }
            });

            console.log(result);
            let value = Object.values(result)[0];
            // console.log(value);
            let tr = '<th scope="col">#</th>';
            for (const key in value[0]) {
                tr += `<td>${key.trim()}</td>`;
                headArr.push(key);
            }
            tableHead.innerHTML = tr;
            for (let i = 1; i <= 20; i++) {
                let tds = `<th scope="row">${i}</th>`
                for (const key in value[i]) {
                    tds += `<td><input class="data-input" type="text" value="${value[i][key]}"></td>`
                }
                tableBody.innerHTML +=
                    `<tr>
                        ${tds}
                    </tr>`
            }

        }
    } catch (e) {
        console.error(e);
    }
}
function save() {
    let inputs = document.querySelectorAll(".data-input");
    let propCount = headArr.length;
    let data = `[{`;
    for (let i = 0; i < inputs.length; i++) {
        data += `"${headArr[i < propCount ? i : i % propCount]}" : "${inputs[i].value}",`;
        if ((i + 1) % propCount === 0 && i !== 0) {
            data = data.slice(0, data.length - 1);
            data += '},\n{';
        }

    }
    data = data.slice(0, data.length - 3);
    data += ']';
    console.log(data);
    try {
        finalData = JSON.parse(data);
    }
    catch (e) {
        alert("Duzgun Parse Olmadi");
        return;
    }
    if (finalData != '') {
        downloadBtn.disabled = false;
    }

}

function download() {
    let filename = 'test.xlsx';
    var ws = XLSX.utils.json_to_sheet(finalData);
    console.log(finalData);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws);
    XLSX.writeFile(wb, filename);
}