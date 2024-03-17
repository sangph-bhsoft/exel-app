import { useRef, useState } from "react";
import saveAs from "file-saver";
import ExcelJS from "exceljs";
import TeenClass from "./components/class_10";
import NewVersion from "./components/new-version";

function App() {
  const [source, setSource] = useState([]);
  const [destination, setDestination] = useState([]);
  const [file, setFile] = useState({
    source: "",
    destination: "",
  });
  const cells = [
    "A",
    "B",
    "C",
    "D",
    "E",
    "F",
    "G",
    "H",
    "I",
    "J",
    "K",
    "L",
    "M",
    "N",
    "O",
    "P",
    "Q",
    "R",
    "S",
    "T",
    "U",
    "V",
    "W",
    "X",
    "Y",
    "Z",
    "AA",
    "AB",
  ];
  const inputFile = useRef(null);
  const outputFile = useRef(null);

  const handleImportInputFile = () => {
    inputFile.current.click();
  };

  const handleImportOutputFile = () => {
    outputFile.current.click();
  };

  const fileInputHandler = (event) => {
    const file = event.target.files[0];
    setFile((f) => ({
      ...f,
      [event.target.name]: file.name,
    }));
    const wb = new ExcelJS.Workbook();
    const reader = new FileReader();

    reader.readAsArrayBuffer(file);
    reader.onload = () => {
      const buffer = reader.result;
      wb.xlsx
        .load(buffer)
        .then((workbook) => {
          const sheet = workbook.getWorksheet(1);
          const rows = [];
          sheet.eachRow((row, rowIndex) => {
            if (row.values[0] === undefined) {
              rows.push(row.values.slice(1));
            } else {
              rows.push(row.values);
            }
          });

          if (event.target.name === "source") {
            setSource(rows);
          } else {
            setDestination(rows);
          }
        })
        .catch((err) => {
          console.log(err);
        });
    };
  };

  const handleDownload = () => {
    let studentSources = [];
    let studentDestinations = [];
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("Sheet1");
    ws.columns = [
      { header: "STT", key: "STT", width: 8 },
      { header: "Mã lớp", key: "Mã lớp", width: 8 },
      {
        header: "Mã định danh Bộ GD&ĐT",
        key: "Mã định danh Bộ GD&ĐT",
        width: 15,
      },
      { header: "Họ tên", key: "Họ tên", width: 30 },
      { header: "Ngày sinh", key: "Ngày sinh", width: 15 },
      { header: "Toán", key: "Toán", width: 10 },
      { header: "Vật lý", key: "Vật lý", width: 10 },
      { header: "Hóa học", key: "Hóa học", width: 10 },
      { header: "Sinh học", key: "Sinh học", width: 10 },
      { header: "Tin học", key: "Tin học", width: 10 },
      { header: "Ngữ văn", key: "Ngữ văn", width: 10 },
      { header: "Lịch sử", key: "Lịch sử", width: 10 },
      { header: "Địa lí", key: "Địa lí", width: 10 },
      { header: "Ngoại ngữ", key: "Ngoại ngữ", width: 10 },
      { header: "Công nghệ", key: "Công nghệ", width: 10 },
      { header: "GD QP-AN", key: "GD QP-AN", width: 10 },
      { header: "Thể dục", key: "Thể dục", width: 10 },
      { header: "Ngoại ngữ 2", key: "Ngoại ngữ 2", width: 10 },
      { header: "Nghề phổ thông", key: "Nghề phổ thông", width: 10 },
      { header: "GDCD", key: "GDCD", width: 10 },
      { header: "ĐTB các môn", key: "ĐTB các môn", width: 10 },
      { header: "Học lực", key: "Học lực", width: 8 },
      { header: "Hạnh kiểm", key: "Hạnh kiểm", width: 8 },
      { header: "Danh hiệu thi đua", key: "Danh hiệu thi đua", width: 20 },
    ];
    const header1 = source[0];
    for (let i = 0; i < header1.length; i++) {
      let k = i;
      for (let j = i + 1; j < header1.length; j++) {
        if (header1[i] === header1[j]) {
          k++;
        }
      }
      if (k > i) {
        ws.mergeCells(`${cells[i]}1 : ${cells[k]}1`);
        i += k - i;
      }
      ws.getCell(`${cells[i]}1`).value = header1[i];
    }
    const header2 = source[1];
    for (let i = 0; i < header1.length; i++) {
      if (header1[i] === header2[i]) {
        ws.mergeCells(`${cells[i]}1 : ${cells[i]}2`);
      }
      ws.getCell(`${cells[i]}2`).value = header2[i];
    }

    const header = destination[3];

    for (let i = 4; i < destination.length; i++) {
      const row = destination[i];
      let student = {};
      for (let j = 0; j < row.length; j++) {
        Object.assign(student, { [header[j]]: row[j] });
      }
      studentDestinations.push(student);
    }

    for (let i = 2; i < source.length; i++) {
      const row = source[i];
      let student = {};
      Object.assign(student, {
        STT: row[0],
        "Mã lớp": row[1],
        "Mã định danh Bộ GD&ĐT": row[2],
        "Họ tên": row[3],
        "Ngày sinh": row[4],
      });
      studentSources.push(student);
    }

    studentSources = studentSources.filter(
      (item) => !!item["Họ tên"] && Number.isInteger(+item["STT"])
    );
    studentDestinations = studentDestinations.filter(
      (item) => !!item["Họ và tên"] && Number.isInteger(+item["STT"])
    );

    const studentRenders = studentSources.map((st) => {
      let students = studentDestinations.filter(
        (std) =>
          std["Họ và tên"].normalize("NFC") === st["Họ tên"].normalize("NFC")
      );
      if (students.length > 1) {
        window.alert(`Có  học sinh trùng tên ${st["Họ và tên"]}`);
      }

      let student = students.length && students[0];

      if (student) {
        return {
          ...st,
          Toán: student["TOÁN"],
          "Vật lý": student["VẬT LÝ"],
          "Hóa học": student["HÓA HỌC"],
          "Sinh học": student["SINH HỌC"],
          "Tin học": student["TIN HỌC"],
          "Ngữ văn": student["NGỮ VĂN"],
          "Lịch sử": student["LỊCH SỬ"],
          "Địa lí": student["ĐỊA LÝ"],
          "Ngoại ngữ": student["NGOẠI NGỮ"],
          "Công nghệ": student["CÔNG_NGHỆ"]
            ? student["CÔNG_NGHỆ"]
            : student["CÔNG NGHỆ"],
          "GD QP-AN": student["QP-AN"],
          "Thể dục": student["THỂ DỤC"],
          "Ngoại ngữ 2": student["NGOẠI NGỮ 2"],
          "Nghề phổ thông": student["NGHỀ PHỔ THÔNG"],
          GDCD: student["GDCD"],
          "ĐTB các môn": student["Trung bình các môn"],
          "Học lực": student["HL"],
          "Hạnh kiểm": student["HK"],
          "Danh hiệu thi đua": student["TĐ"],
        };
      } else {
        window.alert(`Học sinh ${st["Họ tên"]} không có điểm`);
        return st;
      }
    });

    studentRenders.forEach((st) => {
      ws.addRow(st);
    });
    workbook.xlsx.writeBuffer().then(function (buffer) {
      saveAs(
        new Blob([buffer], { type: "application/octet-stream" }),
        `FileHoanThanh-${file.source}`
      );
    });
    setSource([]);
    setDestination([]);
    setFile({
      source: "",
      destination: "",
    });
  };

  return (
    <div className="container mx-auto">
      <div className="flex flex-col space-y-5  w-1/2 mx-auto bg-gray-100 mt-10 rounded-sm shadow-lg px-4 py-4">
        <div className="space-y-2 w-full">
          <div>Nhập file excel mẫu</div>
          <div className="flex items-center rounded-lg overflow-hidden">
            <input
              type="file"
              ref={inputFile}
              className="hidden"
              onChange={fileInputHandler}
              name="source"
            />
            <input
              className="flex-1 px-2 py-2 text-purple-600 italic outline-none focus:outline-none"
              readOnly
              value={file.source}
            />
            <button
              onClick={handleImportInputFile}
              className="bg-orange-400 px-2 py-2 text-white uppercase"
            >
              Import
            </button>
          </div>
        </div>

        <div className="space-y-2 w-full">
          <div>Nhập file excel dữ liệu</div>
          <div className="flex items-center rounded-lg overflow-hidden">
            <input
              type="file"
              ref={outputFile}
              className="hidden"
              onChange={fileInputHandler}
              name="destination"
            />
            <input
              className="flex-1 px-2 py-2 text-purple-600 italic outline-none focus:outline-none"
              readOnly
              value={file.destination}
            />
            <button
              onClick={handleImportOutputFile}
              className="bg-orange-400 px-2 py-2 text-white uppercase"
            >
              Import
            </button>
          </div>
        </div>

        <button
          onClick={handleDownload}
          disabled={!source.length || !destination.length}
          className="bg-pink-500 text-white active:bg-pink-600 font-bold uppercase text-sm px-6 py-3 rounded shadow hover:shadow-lg outline-none focus:outline-none mr-1 mb-1 ease-linear transition-all duration-150"
        >
          Download
        </button>
      </div>

      <TeenClass />

      <NewVersion />
    </div>
  );
}

export default App;
