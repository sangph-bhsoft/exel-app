import React, { useRef, useState } from "react";
import ExcelJS from "exceljs";
import saveAs from "file-saver";

const TeenClass = () => {
  const [source, setSource] = useState([]);
  const [destination, setDestination] = useState([]);
  const [isCN, setIsCN] = useState(false);
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
    "AC",
    "AD",
    "AE",
    "AF",
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
    let columns = [
      { header: "STT", key: "STT", width: 8 },
      { header: "Mã lớp", key: "Mã lớp", width: 8 },
      {
        header: "Mã học sinh",
        key: "Mã học sinh",
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
      { header: "Ngoại ngữ 1", key: "Ngoại ngữ 1", width: 10 },
      { header: "Công nghệ", key: "Công nghệ", width: 10 },
      { header: "GD QP-AN", key: "GD QP-AN", width: 10 },
      { header: "Ngoại ngữ 2", key: "Ngoại ngữ 2", width: 10 },
      { header: "Toán Pháp", key: "Toán Pháp", width: 10 },
      {
        header: "Môn tự chọn song ngữ",
        key: "Môn tự chọn song ngữ",
        width: 10,
      },
      //
      { header: "Giáo dục thể chất", key: "Giáo dục thể chất", width: 10 },
      {
        header: "Hoạt động trải nghiệm",
        key: "Hoạt động trải nghiệm",
        width: 10,
      },
      { header: "Giáo dục địa phương", key: "Giáo dục địa phương", width: 10 },
      { header: "Mỹ thuật", key: "Mỹ thuật", width: 10 },
      { header: "Âm nhạc", key: "Âm nhạc", width: 10 },
      {
        header: "Tiếng dân tộc thiểu số",
        key: "Tiếng dân tộc thiểu số",
        width: 10,
      },
      {
        header: "Giáo dục kinh tế và pháp luật",
        key: "Giáo dục kinh tế và pháp luật",
        width: 10,
      },
      { header: "Kết quả rèn luyện", key: "Kết quả rèn luyện", width: 10 },
      { header: "Kết quả học tập", key: "Kết quả học tập", width: 10 },
    ];

    if (isCN) {
      columns = [
        ...columns,
        { header: "Danh hiệu cả năm", key: "Danh hiệu cả năm", width: 10 },
        {
          header: "TS ngày nghỉ học cả năm",
          key: "TS ngày nghỉ học cả năm",
          width: 10,
        },
        { header: "Được lên lớp", key: "Được lên lớp", width: 10 },
        {
          header: "Kiểm tra lại, rèn luyện HK trong hè",
          key: "Kiểm tra lại, rèn luyện HK trong hè",
          width: 10,
        },
      ];
    }

    ws.columns = columns;
    const header1 = source[0];
    for (let i = 0; i < header1.length; i++) {
      ws.getCell(`${cells[i]}1`).value = header1[i];
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

    for (let i = 1; i < source.length; i++) {
      const row = source[i];
      let student = {};
      Object.assign(student, {
        STT: row[0],
        "Mã lớp": row[1],
        "Mã học sinh": row[2],
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
          "Ngoại ngữ 1": student["NGOẠI NGỮ"],
          "Công nghệ": student["CÔNG NGHỆ"],
          "GD QP-AN": student["QP-AN"],
          "Ngoại ngữ 2": student["NGOẠI NGỮ 2"],
          "Toán Pháp": student["TOAN PHÁP"],
          "Môn tự chọn song ngữ": student["MÔN TỰ CHỢN SONG NGỮ"],
          "Giáo dục thể chất": student["THỂ DỤC"],
          "Hoạt động trải nghiệm": student["HOẠT ĐỘNG NGOÀI GIỜ LÊN LỚP"],
          "Giáo dục địa phương": student["GIÁO DỤC ĐỊA PHƯƠNG"],
          "Mỹ thuật": student["MỸ THUẬT"],
          "Âm nhạc": student["ÂM NHẠC"],
          "Tiếng dân tộc thiểu số": student["TIẾNG DÂN TỘC THIỂU SỐ"],
          "Giáo dục kinh tế và pháp luật": student["GDCD"],
          "Kết quả rèn luyện": student["HK"],
          "Kết quả học tập": student["HL"],
          "Danh hiệu cả năm": student["TD"],
          "TS ngày nghỉ học cả năm": student["TS NGÀY NGHỈ HỌC CẢ NĂM"],

          "Được lên lớp": student["ĐƯƠC LÊN LỚP"],
          "Kiểm tra lại, rèn luyện HK trong hè": "Điểm kiểm tra lại",
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
        <div className="flex items-center justify-between">
          <div className="my-2 text-2xl font-bold">Lớp 10</div>
          <div>
            <input
              type="checkbox"
              value={isCN}
              onChange={(e) => setIsCN(e.target.checked)}
            />
            <span className="ml-2">Cả năm</span>
          </div>
        </div>
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
    </div>
  );
};

export default TeenClass;
