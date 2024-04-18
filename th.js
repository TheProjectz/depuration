document.addEventListener("DOMContentLoaded", () => {
  const file = document.getElementById("fileInput");

  file.addEventListener("change", (e) => {
    document.getElementById("btnStart").removeAttribute("disabled");
    const exc = e.target.files[0];

    const file = new FileReader();

    file.onload = function (c) {
      const data = c.target.result;

      const workbook = XLSX.read(data, {
        type: "binary",
      });
      workbook.SheetNames.forEach((sheet) => {
        const rowObject = XLSX.utils.sheet_to_row_object_array(
          workbook.Sheets[sheet]
        );

        const arryData = Object.values(rowObject);

        let header = [];

        Object.keys(arryData[0]).forEach((key) => {
          header.push(key);
        });

        const table = document.getElementById("tb");

        table.querySelector("tbody").remove();
        table.querySelector("thead>tr").remove();
        table.querySelector("thead").append(document.createElement("tr"));
        table.append(document.createElement("tbody"));

        header.forEach((hed) => {
          table.querySelector("thead>tr").innerHTML += `
                <th scope="col">${hed}</th>
            `;
        });

        for (let i = 0; i < arryData.length; i++) {
          const temp = Object.values(arryData[i]);
          const newRow = document.createElement("tr");
          temp.forEach((dat) => {
            const newCell = document.createElement("td");
            newCell.textContent = dat;
            newRow.appendChild(newCell);
          });
          table.querySelector("tbody").appendChild(newRow);
        }

        const btnStart = document.getElementById("btnStart");
        btnStart.addEventListener("click", ()=>{
            document.getElementById("btnSave").removeAttribute("disabled");
        })

        const btnSave = document.getElementById("btnSave");
        btnSave.addEventListener("click", () => {
          const json = [];
          arryData.forEach((dt) => {
            json.push({ "Codigo": dt.CÓDIGO, "Definitiva": dt.Def  })
          });
          const worksheet = XLSX.utils.json_to_sheet(json);
          //const workbook = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(workbook, worksheet, "DatosNuevos");
          XLSX.utils.sheet_add_aoa(worksheet, [["CÓDIGO", "Def"]], {
            origin: "A1",
          });
          XLSX.writeFile(workbook, "DatosCompletos.xlsx", {
            compression: true,
          });
        });
      });
    };

    file.readAsBinaryString(exc);
  });
});
