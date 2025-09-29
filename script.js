const fileInput = document.getElementById("excel");
const fileInputDiv = document.getElementById("fileInputDiv");
const outputWrapper = document.getElementById("outputWrapper");
const outputDiv = document.getElementById("outputDiv");
const downloadBtn = document.getElementById("downloadBtn");

fileInput.addEventListener("change", function () {

  const file = fileInput.files[0];
  if (!file) return;
  const filename = file.name.toLowerCase();

  if (filename.endsWith(".xls") || filename.endsWith(".xlsx")) {

    fileInputDiv.classList.add("d-none");
    outputWrapper.classList.remove("d-none");
    downloadBtn.classList.remove("d-none");

    const reader = new FileReader();
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetname = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetname];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      const consignee = getValueNextCell("Consignee", rows);
      const address = getValueNextCell("Address", rows);
      const country = getValueNextCell("Country", rows);
      const orderNo = getValueNextCell("Order No", rows)
      const packageNo = getValueNextCell("Package No", rows)

      console.log(rows);
      console.log("Consignee: ", consignee);
      console.log("Address: ", address);
      console.log("Country: ", country);
      console.log("Order No: ", orderNo);
      console.log("Package No: ", packageNo);

      outputDiv.innerHTML = `
        <h4 class="text-center mb-4">Shipping Details</h4>
          <table class="table table-bordered">
            <tbody>
              <tr>
                <th>Consignee:</th>
                <td>${consignee}</td>
              </tr>
              <tr>
                <th>Address:</th>
                <td>${address}</td>
              </tr>
              <tr>
                <th>Country:</th>
                <td>${country}</td>
              </tr>
              <tr>
                <th>Order No:</th>
                <td>${orderNo}</td>
              </tr>
              <tr>
                <th>Package No:</th>
                <td>${packageNo}</td>
              </tr>
            </tbody>
          </table>
        `;

      const keyValueObject = {};
      rows.forEach(row => {
        if (row[0] && row[1]) {
          keyValueObject[row[0].replace(":", "").trim()] = row[1];
        }
      });
      console.log("Key-Value Object:", keyValueObject);
    };
    reader.readAsArrayBuffer(file);
  } else {
    fileInput.value = "";
    new bootstrap.Modal(document.getElementById("warningModal")).show();
  }
});

function getValueNextCell(label, rows) {
  for (row of rows) {
    if (row[0] && row[0].toString().replace(":", "").trim() === label) {
      return row[1] || null;
    }
  }
  return "Not Found!";
}

downloadBtn.addEventListener("click", () => {
  window.print();
});