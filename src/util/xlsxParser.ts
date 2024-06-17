const XLSX = require("xlsx");
const ExcelJS = require("exceljs");

type MAP_DATA_GV = Map<
  string,
  {
    ms_gv: string;
    gv_vnedu: string;
    gv_tkb: string;
  }
>;
type MAP_DATA_MH = Map<
  string,
  {
    mh_vnedu: string;
    mh_tkb: string;
  }
>;
interface DATA_GV {
  STT: number;
  MA_SO_GV: string;
  GV_VNEDU: string;
  GV_TKB: string;
  PERCENT: string;
}
interface DATA_MH {
  STT: number;
  MH_VNEDU: string;
  MH_TKB: string;
  PERCENT: string;
}

function char(n: number) {
  return String.fromCharCode(n + 65).toUpperCase();
}

async function renderData(
  vnedu_path: string,
  tkb_path: string,
  path_data: string
) {
  const vnedu_xlsx = XLSX.readFile(vnedu_path);
  const tkb_xlsx = XLSX.readFile(tkb_path);
  const sheetNamesVNEDU = vnedu_xlsx.SheetNames;
  const sheetNamesTKB = tkb_xlsx.SheetNames;
  const vnedu_sheet = vnedu_xlsx.Sheets[sheetNamesVNEDU[0]];
  const tkb_sheet = tkb_xlsx.Sheets[sheetNamesTKB[0]];
  let row_start_tkb = 6,
    row_start_vnedu = 12;
  const GV_VNEDU: Array<{
    ms_gv: string;
    ten_gv: string;
  }> = [];
  const GV_TKB: Array<string> = [];
  const MH_VNEDU: Array<string> = [];
  const MH_TKB: Array<string> = [];
  // // get GV_VNEDU, MH_VNEDU
  while (
    vnedu_sheet[`${char(0)}${row_start_vnedu}`] ||
    vnedu_sheet[`${char(1)}${row_start_vnedu}`] ||
    vnedu_sheet[`${char(2)}${row_start_vnedu}`] ||
    vnedu_sheet[`${char(3)}${row_start_vnedu}`] ||
    vnedu_sheet[`${char(4)}${row_start_vnedu}`]
  ) {
    const ms_gv = vnedu_sheet[`${char(0)}${row_start_vnedu}`]?.v;
    const ten_gv = vnedu_sheet[`${char(1)}${row_start_vnedu}`]?.v;
    const mh = vnedu_sheet[`${char(2)}${row_start_vnedu}`]?.v;
    if (ms_gv && ten_gv && !GV_VNEDU.find((gv) => gv.ms_gv === ms_gv))
      GV_VNEDU.push({ ms_gv, ten_gv });
    if (mh && !MH_VNEDU.includes(mh)) MH_VNEDU.push(mh);
    row_start_vnedu++;
  }
  // get GV_TKB, MH_TKB
  while (
    tkb_sheet[`${char(0)}${row_start_tkb}`] ||
    tkb_sheet[`${char(1)}${row_start_tkb}`] ||
    tkb_sheet[`${char(2)}${row_start_tkb}`] ||
    tkb_sheet[`${char(3)}${row_start_tkb}`]
  ) {
    const gv = tkb_sheet[`${char(0)}${row_start_tkb}`]?.v;
    const mh = tkb_sheet[`${char(2)}${row_start_tkb}`]?.v;
    if (gv && !GV_TKB.includes(gv)) GV_TKB.push(gv);
    if (mh && !MH_TKB.includes(mh)) MH_TKB.push(mh);
    row_start_tkb++;
  }
  // create file xlsx
  const wb = XLSX.utils.book_new();
  const data_gv: DATA_GV[] = [];
  const data_mh: DATA_MH[] = [];
  GV_VNEDU.forEach((gv) => {
    let data_check = {
      check: false,
      percent: 0,
    };
    let data_add;
    GV_TKB.forEach((gv_tkb) => {
      const check = isSame(gv.ten_gv.split("/")[0], gv_tkb);
      if (check.check && check.percent > data_check.percent) {
        data_check = check;
        data_add = {
          STT: data_gv.length + 1,
          MA_SO_GV: gv.ms_gv,
          GV_VNEDU: gv.ten_gv,
          GV_TKB: gv_tkb,
          PERCENT: data_check.percent.toString(),
        };
      }
    });
    if (data_check.check) data_gv.push(data_add);
    else {
      data_add = {
        STT: data_gv.length + 1,
        MA_SO_GV: gv.ms_gv,
        GV_VNEDU: gv.ten_gv,
        GV_TKB: "",
        PERCENT: "",
      };
      data_gv.push(data_add);
    }
  });
  MH_TKB.forEach((mh_tkb) => {
    let data_check = {
      check: false,
      percent: 0,
    };
    let data_add;
    MH_VNEDU.forEach((mh) => {
      const check = isSameSubject(mh, mh_tkb);
      if (check.check && check.percent > data_check.percent) {
        data_check = check;
        data_add = {
          STT: data_mh.length + 1,
          MH_TKB: mh_tkb,
          MH_VNEDU: mh,
          PERCENT: data_check.percent.toString(),
        };
      }
    });
    if (data_check.check) data_mh.push(data_add);
    else {
      data_add = {
        STT: data_mh.length + 1,
        MH_TKB: mh_tkb,
        MH_VNEDU: "",
        PERCENT: "",
      };
      data_mh.push(data_add);
    }
  });
  const sheet_gv = XLSX.utils.json_to_sheet(data_gv);
  const sheet_mh = XLSX.utils.json_to_sheet(data_mh);
  XLSX.utils.book_append_sheet(wb, sheet_gv, "DATA_GV");
  XLSX.utils.book_append_sheet(wb, sheet_mh, "DATA_MH");
  XLSX.writeFile(wb, path_data);
  return null;
}

const convertTKBToVNEDU = async (
  path_tkb: string,
  path_data: string,
  path_result: string,
  is_hk2?: boolean
) => {
  const work_book_tkb = new ExcelJS.Workbook();
  const work_book_data = new ExcelJS.Workbook();
  const work_book_result = new ExcelJS.Workbook();
  const tkb_xlsx = await work_book_tkb.xlsx.readFile(path_tkb);
  const data_xlsx = await work_book_data.xlsx.readFile(path_data);
  const result_xlsx = await work_book_result.xlsx.readFile(path_result);
  let sheet_data_gv, sheet_data_mh;
  data_xlsx.eachSheet((sheet) => {
    switch (sheet.name) {
      case "DATA_GV":
        sheet_data_gv = sheet;
        break;
      case "DATA_MH":
        sheet_data_mh = sheet;
        break;
    }
  });
  const sheet_result = result_xlsx.getWorksheet(1);
  const sheet_tkb = tkb_xlsx.getWorksheet(1);
  // create map data
  const map_data_gv: MAP_DATA_GV = new Map();
  const map_data_mh: MAP_DATA_MH = new Map();
  let row_data_gv = 2;
  while (sheet_data_gv.getCell(`${char(0)}${row_data_gv}`)?.value) {
    const ms_gv_vnedu =
      sheet_data_gv.getCell(`${char(1)}${row_data_gv}`)?.value ?? "";
    const gv_vnedu =
      sheet_data_gv.getCell(`${char(2)}${row_data_gv}`)?.value ?? "";
    const gv_tkb = sheet_data_gv.getCell(`${char(3)}${row_data_gv}`)?.value;
    if (gv_tkb)
      map_data_gv.set(gv_tkb, {
        ms_gv: ms_gv_vnedu,
        gv_vnedu: gv_vnedu,
        gv_tkb: gv_tkb,
      });
    row_data_gv++;
  }
  let row_data_mh = 2;
  while (sheet_data_mh.getCell(`${char(0)}${row_data_mh}`)?.value) {
    const mh_tkb = sheet_data_mh.getCell(`${char(1)}${row_data_mh}`)?.value;
    const mh_vnedu =
      sheet_data_mh.getCell(`${char(2)}${row_data_mh}`)?.value ?? "";
    if (mh_tkb)
      map_data_mh.set(mh_tkb, {
        mh_vnedu: mh_vnedu,
        mh_tkb: mh_tkb,
      });
    row_data_mh++;
  }
  // get GV_TKB, MH_TKB
  let row_start_tkb = 6,
    row_start_vnedu = 12,
    teacher_current = {
      ms_gv: undefined,
      gv_vnedu: undefined,
    };
  while (
    sheet_tkb.getCell(`${char(0)}${row_start_tkb}`)?.value ||
    sheet_tkb.getCell(`${char(1)}${row_start_tkb}`)?.value ||
    sheet_tkb.getCell(`${char(2)}${row_start_tkb}`)?.value ||
    sheet_tkb.getCell(`${char(3)}${row_start_tkb}`)?.value
  ) {
    const gv_tkb = sheet_tkb.getCell(`${char(0)}${row_start_tkb}`).value ?? "";
    const mh_tkb = sheet_tkb.getCell(`${char(2)}${row_start_tkb}`).value ?? "";
    const classes = sheet_tkb.getCell(`${char(3)}${row_start_tkb}`).value ?? "";
    //
    const ms_gv_vnedu = map_data_gv.has(gv_tkb)
      ? map_data_gv.get(gv_tkb).ms_gv
      : "";
    const gv_vnedu = map_data_gv.has(gv_tkb)
      ? map_data_gv.get(gv_tkb).gv_vnedu
      : "";
    const mh_vnedu = map_data_mh.has(mh_tkb)
      ? map_data_mh.get(mh_tkb).mh_vnedu
      : "";
    // add row
    sheet_result.getRow(row_start_vnedu).values = [
      ms_gv_vnedu !== teacher_current?.ms_gv ? ms_gv_vnedu : "",
      gv_vnedu !== teacher_current?.gv_vnedu ? gv_vnedu : "",
      mh_vnedu,
      is_hk2 ? "" : formatClass(classes),
      is_hk2 ? formatClass(classes) : "",
    ];
    for (let j = 0; j < 5; j++) {
      sheet_result.getCell(`${char(j)}${row_start_vnedu}`).font = {
        name: "Times New Roman", // Đặt font family
        size: 13, // Đặt kích thước font
      };
    }

    teacher_current = {
      ms_gv: ms_gv_vnedu,
      gv_vnedu: gv_vnedu,
    };
    row_start_tkb++;
    row_start_vnedu++;
  }
  await result_xlsx.xlsx.writeFile(path_result);
  return;
};

export const xlsxParser = {
  renderData,
  convertTKBToVNEDU,
};

const isSame = (a: string, b: string): { check: boolean; percent: number } => {
  let check = false;
  let percent = 0;
  if (a.toLowerCase() === b.toLowerCase()) {
    check = true;
    percent = 1;
  }
  if (!check && (b.includes(a) || a.includes(b))) {
    check = true;
    percent = 0.8;
  }
  // check word
  if (!check) {
    const _array_a = a.split(" ");
    const _array_b = b.split(" ");
    if (_array_a.length === _array_b.length) {
      let count = 0;
      for (let i = 0; i < _array_a.length; i++) {
        if (_array_a[i].toLowerCase() === _array_b[i].toLowerCase()) {
          count++;
        }
      }
      if (count / _array_a.length > 0.7) {
        check = true;
        percent = count / _array_a.length;
      }
    }
  }
  return {
    check: check,
    percent: percent,
  };
};

const isSameSubject = (
  a: string,
  b: string
): { check: boolean; percent: number } => {
  let check = false;
  let percent = 0;
  if (a.toLowerCase() === b.toLowerCase()) {
    check = true;
    percent = 1;
  }
  if (!check && (b.includes(a) || a.includes(b))) {
    check = true;
    percent = 0.8;
  }
  // check word
  if (!check) {
    const _array_a = a.split(" ");
    const _array_b = b.split(" ");
    if (_array_a.length === _array_b.length) {
      let count = 0;
      for (let i = 0; i < _array_a.length; i++) {
        if (_array_a[i].toLowerCase() === _array_b[i].toLowerCase()) {
          count++;
        }
      }
      if (count / _array_a.length > 0.7) {
        check = true;
        percent = count / _array_a.length;
      }
    }
  }
  if (!check) {
    let _a = (a as any).replaceAll(" ", "");
    let _b = (b as any).replaceAll(" ", "");
    let _array_a;
    if (_a.toUpperCase() === _a) {
      _array_a = _a.split("");
    } else {
      _array_a = a.split(" ").map((item) => item[0].toUpperCase());
    }
    let _array_b;
    if (_b.toUpperCase() === _b) {
      _array_b = _b.split("");
    } else {
      _array_b = b.split(" ").map((item) => item[0].toUpperCase());
    }
    let count = 0;
    let max_length = Math.max(_array_a.length, _array_b.length);
    for (let i = 0; i < max_length; i++) {
      if (_array_a?.[i] === _array_b?.[i]) {
        count++;
      }
    }
    if (count / max_length > 0.6) {
      check = true;
      percent = count / max_length;
    }
  }
  return {
    check: check,
    percent: percent,
  };
};
function formatClass(str: string): string {
  let result = (str as any).replaceAll("  ", " ");
  while (result.includes("(") && result.length > 0) {
    const index_start = result.indexOf("(");
    const index_end =
      result.indexOf(")") > -1 ? result.indexOf(")") : index_start + 1;
    result = result.slice(0, index_start) + result.slice(index_end + 1);
  }
  return result;
}
