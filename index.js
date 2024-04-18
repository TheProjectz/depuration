class Debuging {
  constructor(exc) {
    this.exc = exc.target.files[0];
    this.arrData = new arrDataWorBook();
    this.WorBook = new arrDataWorBook();
  }

  getExc() {
    return this.exc;
  }

  setArrayData(rowObject) {
    this.arrData.setArryData(rowObject);
  }

  getArrayData() {
    return this.arrData.getArryData();
  }

  setWorkBook(data) {
    this.WorBook.setWorkBook(data);
  }

  getWorkBook() {
    return this.WorBook.getWorkbook();
  }

  onLoaded(file) {
    return (file.onload = (c) => {
      const data = c.target.result;
      this.setWorkBook(data);
      const workbook = this.getWorkBook();

      //aqui se manda el parametro para mirar que hoja se quiere visualizar
      //const sheetName = "Hoja1";
      workbook.SheetNames.forEach((sheetName) => {
        const rowObject = XLSX.utils.sheet_to_row_object_array(
          workbook.Sheets[sheetName]
        );
        //
        const arryData = Object.values(rowObject);

        this.setArrayData(rowObject);

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
      });
    });
  }
}

class arrDataWorBook {
  const = (this.rowObject = null);
  const = (this.workbook = null);

  setArryData(rowObject) {
    this.rowObject = Object.values(rowObject);
  }

  getArryData() {
    return this.rowObject;
  }

  setWorkBook(data) {
    this.workbook = XLSX.read(data, {
      type: "binary",
    });
  }

  getWorkbook() {
    return this.workbook;
  }
}

document.addEventListener("DOMContentLoaded", () => {
  const file = document.getElementById("fileInput");

  file.addEventListener("change", (e) => {
    document.getElementById("btnStart").removeAttribute("disabled");
    document.getElementById("btnSave").setAttribute("disabled", "disabled");
    const exc = new Debuging(e);

    const file = new FileReader();
    exc.onLoaded(file);
    file.readAsBinaryString(exc.getExc());

    const btnStart = document.getElementById("btnStart");
    btnStart.addEventListener("click", () => {
      document.getElementById("btnSave").removeAttribute("disabled");
    });

    const btnSave = document.getElementById("btnSave");
    btnSave.addEventListener("click", () => {
      const arrDt = exc.getArrayData();
      const json = [];
      arrDt.forEach((dt) => {
        json.push({ Codigo: dt.CÓDIGO, Definitiva: dt.Def });
      });
      const worksheet = XLSX.utils.json_to_sheet(json);
      //const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(exc.getWorkBook(), worksheet, "DatosNuevos");
      XLSX.utils.sheet_add_aoa(worksheet, [["CÓDIGO", "Def"]], {
        origin: "A1",
      });
      XLSX.writeFile(exc.getWorkBook(), "DatosCompletos.xlsx", {
        compression: true,
      });
    });
  });
});
