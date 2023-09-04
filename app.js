const regexPattern = /^\d{2}-[A-Za-z]{3}-\d{2} \d{2}:\d{2}:\d{2}$/;
const input = document.getElementById("input");
const btn = document.querySelector(".btn");
btn.addEventListener("click", () => {
  readXlsxFile(input.files[0]).then((rows) => {
    checkerPayDirect(rows, 2);
  });
});

const inputWeb = document.getElementById("inputweb");
const btnWeb = document.querySelector(".btnWeb");

btnWeb.addEventListener("click", () => {
  readXlsxFile(inputWeb.files[0]).then((rows) => {
    checkerWeb(rows, 1);
  });
});

const banks = [
  ["Zenith", "ZIB", "ZBI"],
  ["GTB", "GTI"],
  ["Access", "ABP"],
  ["ECO"],
  ["Fidelity", "FBP", "FDB"],
  ["FBN"],
  ["FCMB"],
  ["Heritage", "HBP"],
  ["Keystone", "KSB"],
  ["Polaris", "SKYE"],
  ["STANBIC", "SIB"],
  ["Sterlin", "SBP"],
  ["UBN"],
  ["UBA"],
  ["UNITY"],
  ["WEMA", "QPT"],
];

function checkerPayDirect(rows, idx) {
  const hash = {};
  for (let i = 0; i < rows.length; i++) {
    if (regexPattern.test(rows[i][idx])) {
      let bank = rows[i][1].split("|")[0];
      let proceed = false;
      for (let j = 0; j < banks.length; j++) {
        if (banks[j].includes(bank)) {
          bank = banks[j][0];
          proceed = true;
          break;
        }
      }
      if (proceed) {
        const data = {
          date: rows[i][idx].slice(0, 9),
          value: Number(rows[i][8].slice(2).replace(",", "")),
          bank: bank,
        };
        if (!hash.hasOwnProperty(data.bank)) {
          hash[data.bank] = {};
        }
        if (hash[data.bank].hasOwnProperty(data.date)) {
          hash[data.bank][data.date].value += data.value;
        } else {
          hash[data.bank][data.date] = { value: data.value };
        }
      }
    }
  }

  const t = [];
  for (let i in hash) {
    for (let x in hash[i]) {
      t.push([x, hash[i][x].value, "", "", i]);
    }
  }
  generateExcel(t);
}

function checkerWeb(rows, idx) {
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
    ["Date", "Paydirect", "Bank Statement", "Difference", "bank"],
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
