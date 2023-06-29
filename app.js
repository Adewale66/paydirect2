// File.
const regexPattern = /^\d{2}-[A-Za-z]{3}-\d{2} \d{2}:\d{2}:\d{2}$/;

const input = document.getElementById("input");
const btn = document.querySelector(".btn");
btn.addEventListener("click", () => {
  readXlsxFile(input.files[0]).then((rows) => {
    checker(rows, 2);
  });
});

const input2 = document.getElementById("inputweb");
const btn2 = document.querySelector(".btnWeb");

btn2.addEventListener("click", () => {
  readXlsxFile(input2.files[0]).then((rows) => {
    checker(rows, 1);
  });
});

function checker(rows, idx) {
  const hash = {};
  for (let i = 0; i < rows.length; i++) {
    if (regexPattern.test(rows[i][idx])) {
      const data = {
        key: rows[i][idx].slice(0, 9),
        value: Number(rows[i][8].slice(2).replace(",", "")),
      };
      if (data.key in hash) hash[data.key] += data.value;
      else hash[data.key] = data.value;
    }
  }
  const t = [];
  for (let i in hash) t.push([i, hash[i], "", ""]);

  generateExcel(t);
}

function generateExcel(items) {
  const workbook = XLSX.utils.book_new();

  const worksheetData = [
    ["Date", "Paydirect", "Bank Statement", "Difference"],
    ...items,
  ];
  const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet 1");

  const excelBuffer = XLSX.write(workbook, { bookType: "xlsx", type: "array" });

  const blob = new Blob([excelBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  const downloadLink = document.createElement("a");
  downloadLink.href = URL.createObjectURL(blob);
  downloadLink.download = "data.xlsx";

  document.body.appendChild(downloadLink);
  downloadLink.click();

  document.body.removeChild(downloadLink);
}
