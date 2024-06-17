const XLSX = require("xlsx");

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
  MH_VNEDU.forEach((mh) => {
    let data_check = {
      check: false,
      percent: 0,
    };
    let data_add;
    MH_TKB.forEach((mh_tkb) => {
      const check = isSameSubject(mh, mh_tkb);
      if (check.check && check.percent > data_check.percent) {
        data_check = check;
        data_add = {
          STT: data_mh.length + 1,
          MH_VNEDU: mh,
          MH_TKB: mh_tkb,
          PERCENT: data_check.percent.toString(),
        };
      }
    });
    if (data_check.check) data_mh.push(data_add);
    else {
      data_add = {
        STT: data_mh.length + 1,
        MH_VNEDU: mh,
        MH_TKB: "",
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

const convertTKBToVNEDU = (
  path_tkb: string,
  path_data: string,
  path_result: string,
  is_hk2?: boolean
) => {
  const tkb_xlsx = XLSX.readFile(path_tkb);
  const data_xlsx = XLSX.readFile(path_data);
  const result_xlsx = XLSX.readFile(path_result);
  let sheet_data_gv, sheet_data_mh;
  for (let i = 0; i < data_xlsx.SheetNames.length; i++) {
    switch (data_xlsx.SheetNames[i]) {
      case "DATA_GV":
        sheet_data_gv = data_xlsx.Sheets[data_xlsx.SheetNames[i]];
        break;
      case "DATA_MH":
        sheet_data_mh = data_xlsx.Sheets[data_xlsx.SheetNames[i]];
        break;
    }
  }
  const sheet_result = result_xlsx.Sheets[result_xlsx.SheetNames[0]];
  const sheet_tkb = tkb_xlsx.Sheets[tkb_xlsx.SheetNames[0]];
  const data_gv = XLSX.utils.sheet_to_json(sheet_data_gv);
  const data_mh = XLSX.utils.sheet_to_json(sheet_data_mh);
  // create map data
  const map_data_gv: MAP_DATA_GV = new Map();
  const map_data_mh: MAP_DATA_MH = new Map();
  data_gv.forEach((gv: DATA_GV) => {
    map_data_gv.set(gv.GV_TKB, {
      ms_gv: gv.MA_SO_GV,
      gv_vnedu: gv.GV_VNEDU,
      gv_tkb: gv.GV_TKB,
    });
  });
  data_mh.forEach((mh: DATA_MH) => {
    map_data_mh.set(mh.MH_TKB, {
      mh_vnedu: mh.MH_VNEDU,
      mh_tkb: mh.MH_TKB,
    });
  });
  // get GV_TKB, MH_TKB
  let row_start_tkb = 6,
    row_start_vnedu = 12;
  while (
    sheet_tkb[`${char(0)}${row_start_tkb}`] ||
    sheet_tkb[`${char(1)}${row_start_tkb}`] ||
    sheet_tkb[`${char(2)}${row_start_tkb}`] ||
    sheet_tkb[`${char(3)}${row_start_tkb}`]
  ) {
    const gv_tkb = sheet_tkb[`${char(0)}${row_start_tkb}`]?.v ?? "";
    const mh_tkb = sheet_tkb[`${char(2)}${row_start_tkb}`]?.v ?? "";
    const classes = sheet_tkb[`${char(3)}${row_start_tkb}`]?.v ?? "";
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
    sheet_result[`${char(0)}${row_start_vnedu}`] = {
      t: "s",
      v: ms_gv_vnedu,
    };
    sheet_result[`${char(1)}${row_start_vnedu}`] = {
      t: "s",
      v: gv_vnedu,
    };
    sheet_result[`${char(2)}${row_start_vnedu}`] = {
      t: "s",
      v: mh_vnedu,
    };
    if (!is_hk2)
      sheet_result[`${char(3)}${row_start_vnedu}`] = {
        t: "s",
        v: formatClass(classes),
      };
    else
      sheet_result[`${char(4)}${row_start_vnedu}`] = {
        t: "s",
        v: formatClass(classes),
      };
    row_start_tkb++;
    row_start_vnedu++;
  }
  XLSX.writeFile(result_xlsx, path_result);
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
