window.addEventListener("load", async function () {
  let workbook = XLSX.read(await (await fetch("./product.xlsx")).arrayBuffer());
  let worksheet = workbook.SheetNames;
  worksheet.forEach((name) => {
    product = XLSX.utils.sheet_to_json(workbook.Sheets[name]);
    DP = product.filter((item) => item.WhCode == "DP");
    TT = product.filter((item) => item.WhCode == "TT");
    TP = product.filter((item) => item.WhCode == "TP");
    CNDP = product.filter((item) => item.WhCode == "CNDP");
  });
});

function createAutoSuggest(element, list, id) {
  htmlText = `<datalist id="${id}">`;
  list.forEach((ele) => {
    htmlText += `<option value="${ele}"> </option>`;
  });
  htmlText += `</datalist>`;
  element.insertAdjacentHTML("beforeend", htmlText);
}

function selectkho(e) {
  var idData = document.querySelector(
    "#warehouse" + " option[value='" + e.value + "']"
  );

  var row = e.parentNode.parentNode;
  var id = row.cells[0].innerHTML;
  var skuCode = row.cells[2];
  if (idData != null) {
    khoSelected = idData.dataset.value;
  }
  if (skuCode.children[1] != null) {
    skuCode.children[1].remove();
  }
  skuCode.childNodes[1].setAttribute("list", id);

  row.cells[2].childNodes[1].value = "";
  row.cells[5].innerHTML = "";
  row.cells[6].innerHTML = "";
  row.cells[7].innerHTML = "";

  switch (khoSelected) {
    case "DP":
      createAutoSuggest(
        skuCode,
        DP.map((c) => c.SKU),
        id
      );
      break;
    case "TT":
      createAutoSuggest(
        skuCode,
        TT.map((c) => c.SKU),
        id
      );
      break;
    case "TP":
      createAutoSuggest(
        skuCode,
        TP.map((c) => c.SKU),
        id
      );
      break;
    case "CNDP":
      createAutoSuggest(
        skuCode,
        CNDP.map((c) => c.SKU),
        id
      );
    default:
      break;
  }
}

function formatVND(number) {
  if (isNaN(number)) {
    return 0;
  }
  return number.toLocaleString("vi-VN", {
    style: "currency",
    currency: "VND",
  });
}

function selectPro(e) {
  var des = product.filter((d) => d.SKU == e.value);
  var row = e.parentNode.parentNode;
  if (des.length > 0) {
    row.cells[5].innerHTML = formatVND(des[0].Price_1);
    row.cells[6].innerHTML = formatVND(des[0].Price_2);
    row.cells[7].innerHTML = des[0].Description;
  } else {
    row.cells[5].innerHTML = "";
    row.cells[6].innerHTML = "";
    row.cells[7].innerHTML = "";
  }
}

function addRow(tableID) {
  var table = document.getElementById(tableID);
  var rowCount = table.rows.length;
  var row = table.insertRow(rowCount);
  row.classList.add("d-flex");

  var colCount = table.rows[0].cells.length;
  for (var i = 0; i < colCount; i++) {
    var newcell = row.insertCell(i);
    if (table.rows[1].cells[i].classList.value) {
      newcell.classList.add(table.rows[1].cells[i].classList.value);
    }
    newcell.style.width = table.rows[1].cells[i].style.width;
    newcell.style.display = table.rows[1].cells[i].style.display;
    if (i == 0) {
      newcell.innerHTML = rowCount;
    } else {
      newcell.innerHTML = table.rows[1].cells[i].innerHTML;
    }
    if (i > 4 && i < 8) {
      newcell.innerHTML = "";
    }
    if (newcell.childNodes[0]) {
      switch (newcell.childNodes[0].type) {
        case "text":
          newcell.childNodes[0].value = "";
          break;
        case "checkbox":
          newcell.childNodes[0].checked = false;
          break;
        case "select-one":
          newcell.childNodes[0].selectedIndex = 0;
          break;
      }
    }
  }
}

function deleteRow(e) {
  var row = e.parentNode.parentNode;
  if (row.cells[0].innerHTML != 1) {
    row.parentNode.removeChild(row);
  }
}

function deleteTable() {
  console.log("delete all");
}

function createRowTable(element, list) {
  tableText = `<tr class="text-center">`;
  list.forEach((ele) => {
    tableText += `<td> ${ele} </td>`;
  });
  tableText += `</tr>`;
  element.insertAdjacentHTML("beforeend", tableText);
}
function formatDate(date) {
  var res = date.split("-");
  return `${res[2]}-${res[1]}-${res[0]}`;
}

function updateQuotation() {
  window.scrollTo(0, 2800);
  let user = document.getElementById("userName").value;
  let company = document.getElementById("company").value;
  let address = document.getElementById("address").value;
  let mail = document.getElementById("mail").value;
  let phone = document.getElementById("phone").value;
  let time = document.getElementById("datePicker").value;

  document.getElementById("quotation_user").innerHTML = user;
  document.getElementById("quotation_company").innerHTML = company;
  document.getElementById("quotation_address").innerHTML = address;
  document.getElementById("quotation_mail").innerHTML = mail;
  document.getElementById("quotation_phone").innerHTML = phone;
  document.getElementById("quotation_datetime").innerHTML = formatDate(time);

  let product = document.getElementById("dataTable").childNodes[5].rows;
  let productQuotation = [];
  for (let i = 0; i < product.length; i++) {
    let item = [];
    item.push(i + 1);
    item.push(product[i].children[7].innerText);
    item.push(product[i].children[2].children[0].value);
    item.push("China/Asia");
    item.push("pcs");
    item.push(product[i].children[3].children[0].value);
    item.push(product[i].children[4].children[0].value);
    item.push(item[5] * item[6]);
    productQuotation.push(item);
  }
  document.getElementById("quotation_row").innerHTML = "";
  document.getElementById("quotation_sum").innerHTML = "";
  sum = 0;
  productQuotation.forEach((list) => {
    sum += list[7];
    createRowTable(document.getElementById("quotation_row"), list);
  });
  document.getElementById("quotation_prosum").innerText = formatVND(sum);
  document.getElementById("quotation_vat").innerHTML =
    document.getElementById("VAT").value;
  document.getElementById("quotation_vatprice").innerHTML = formatVND(
    (document.getElementById("VAT").value * sum) / 100
  );

  document.getElementById("quotation_sum").innerHTML = formatVND(
    sum + (document.getElementById("VAT").value * sum) / 100
  );
}

function printPdf() {
  const ele = document.getElementById("template");
  var opt = {
    jsPDF: {
      unit: "in",
      scale: 0.2,
      format: "letter",
      orientation: "portrait",
    },
  };
  html2pdf().from(ele).toImg().save();
}

function backForm() {
  window.scrollTo(0, 0);
}
