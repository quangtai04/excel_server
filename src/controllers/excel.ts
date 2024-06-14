const { GoogleSpreadsheet } = require("google-spreadsheet");
const { JWT } = require("google-auth-library");
import { handleError, handleSuccess } from "../helpers/response";

const GOOGLE_SERVICE_ACCOUNT_EMAIL =
  "download-excel@excelserver.iam.gserviceaccount.com";
const GOOGLE_PRIVATE_KEY =
  "-----BEGIN PRIVATE KEY-----\nMIIEugIBADANBgkqhkiG9w0BAQEFAASCBKQwggSgAgEAAoIBAQDKgqsPvYzO6FwB\ntp4N85nmXABGBm6VazAxJzJZK9/F+7ry7COwe18qGjuX/uBFk5r7mpAWQKVlRJnP\nKTnDPUwD99hN/nVG5K+Qv7TC9R916DJ/nxCD3aaKHHEFPvEHf9K4XcdVaXFKsB6s\n/WJurOJkP2rV3IzPDo0HqNobwx2bPyjXy6zt7CRIpMgNHBVLvzDudgXc/L/A8iXN\nhdZhiZ89FDqEDXcCua7zevj6bT2VAlLtihUNu4/Ui3gQ6m0/m6RLfnKmMxhcNt11\nHw6BgI8ZKuYOjK8Jm/VE/4KuJWOKX5TTpme+s3o1UUfEswXCbYID89AtOOSiYya9\nu5Yi0q4bAgMBAAECggEAMEFXwr3vriQXPH9IBVoNS6OTmxwQQMGUb7n/2NjID1TP\nNPCJBpY3VICAv9THmzyzew3XFL5dyxZAMmmH7pqOIQnvfJJMXtLCdRMBX01qrD1i\nvx9nn4xzEUj6s33OkHNogm8yPwuLp/j0rlMIoAfJQIsOCZzu3q4AdBlLs77YMLQ9\nH8IfBt+OY+dd2KPA58k1J2LdXo8MAf2Sqw3Ks5wOSZ7Wvewa+Gk/9n/X+cawUSR/\nI44yJnM5+33Udez3ol2Bq/IV39jQFZCW6S4Tka0iqoJfpH05WnTjEiy8nsCGs93y\nIlcOulExesJPkuufmjDb0z9YmNGg3KQ1Vn5Pp4e5mQKBgQDnkBYI68PhWCmpQh1b\nFjtEltx6+JHn3bOPTp8+ksP+Cl4Y6DEsQD2WdKY9+zn5XYrpHXLmuHdu2FL56zvn\nQJZgwdkUwwhPN7veuEBWGcsrHdAP0Gyl7nkU9oI028l30beevWH+5qe2+Gx0agYN\n0zfXkSuYe2CifZbHgTld+69S2QKBgQDf4bNZY1EVbuxqdFDa66OHnUcOsUwjZqpJ\nZ+slYwJJAFV9NUrje+r2iqdIEG0wd0mCMDJdgm4+TObIddYJnh6lzDBVHC26t/9a\nFThHaawfeTX3e0f2jfCqC4cNja/N2NMj5G37GoVA/zXy3aTFi65AeYAZdqVv9zmQ\nyfG/AwbIEwKBgHTnOxiZ3jQfzDiVFjjsClPgTcMPRqnmNUZ+DMsMUUIpfcPZRSnv\n7KfOkDbuZCBOZ1i081MjgbhGIe3mIkHnS4PhmXRv2fwUSRZxsplFQFquHGI/ePp0\nYFCC+s0wwI9rIuZS+ew0CivHUwmalR/ZqHF96qJ6dxjRipOB27Jk4+hRAoGAERue\nKOZ9+7VOO5RH2XLIPES4eVbzCoF94b4fKew28H0mCztXTaraeZx+y/L1ZQ64f0pP\njvW4luopeIgIoxOCJAlGaDqPBAWrbabRdiONE5qflRnGlgCis1vOJir9lC1NdSCv\nhtCv/heCd9yYCsBxwFuIfmmimru5mmbUIlSI4CECf2GLo/ZUiU/Om54yMxNwxlFa\nexQPw3owo/YEP1xbBdjIBXv9bI2N022orsgqa6/U487rFPG+MTL2yzIZoPuWuVsn\nplgMb9Dio40JnhQhe2jB7f3g+LFfYqecUXd15Rag34Qo4Q5aNQvU8RJEaPrfRyc7\nceEAbAOk/ugE4fz9Ggo=\n-----END PRIVATE KEY-----\n";
const serviceAccountAuth = new JWT({
  // env var values here are copied from service account credentials generated by google
  // see "Authentication" section in docs for more info
  email: GOOGLE_SERVICE_ACCOUNT_EMAIL,
  key: GOOGLE_PRIVATE_KEY,
  scopes: [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
  ],
});
const API = {
  count_api: 0,
  MAX_API: 50,
};
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
function funcSortData(
  index: number,
  sheet_data: any,
  data_vnedu: any[],
  data_tkb: any[],
  type: "GV" | "MH" = "GV"
): Promise<void> {
  if (index >= data_vnedu.length) return Promise.resolve();
  return new Promise(async (resolve, reject) => {
    const [id_vnedu, value, ms_gv = undefined] = data_vnedu[index]._rawData;
    const [name_vnedu, account] = value.split("/");
    let data_add: Array<any>;
    for (let i = 0; i < data_tkb.length; i++) {
      const [id_tkb, name_tkb] = data_tkb[i]._rawData;
      let data_check;
      if (type === "GV") {
        data_check = isSame(name_vnedu, name_tkb);
      } else {
        data_check = isSameSubject(name_vnedu, name_tkb);
      }
      if (data_check?.check && (data_add?.[3] ?? 0) < data_check?.percent) {
        data_add = [id_vnedu, name_vnedu, name_tkb, data_check?.percent];
      }
    }
    if (API.count_api > API.MAX_API) {
      API.count_api = 0;
      await new Promise((resolve) => setTimeout(resolve, 5000));
    }
    API.count_api++;
    sheet_data
      .addRow({
        [`STT_${type}_VNEDU`]: id_vnedu,
        [`${type}_VNEDU`]: name_vnedu,
        [`${type}_TKB`]: data_add?.[2] ?? "",
        [`PERCENT`]: data_add?.[3] ?? "",
        ...(type === "GV" ? { MA_SO_GV: ms_gv } : {}),
      })
      .then(() => {
        funcSortData(index + 1, sheet_data, data_vnedu, data_tkb, type).then(
          () => {
            resolve();
          }
        );
      });
  });
}
export const sortData = async (req, res) => {
  const { sheet_id, stt_start, type_sort } = req.body;
  let type = type_sort ?? "GV";
  try {
    const doc = new GoogleSpreadsheet(sheet_id, serviceAccountAuth);
    await doc.loadInfo();
    let sheet_data, sheet_vnedu, sheet_tkb;
    for (let i = 0; i < doc.sheetCount; i++) {
      switch (doc.sheetsByIndex[i]?.title?.toUpperCase()) {
        case `DATA_${type}`:
          sheet_data = doc.sheetsByIndex[i];
          break;
        case `${type}_VNEDU`:
          sheet_vnedu = doc.sheetsByIndex[i];
          break;
        case `${type}_TKB`:
          sheet_tkb = doc.sheetsByIndex[i];
          break;
      }
    }
    // get data vnedu
    const data_vnedu = await sheet_vnedu.getRows();
    // get data tkb
    const data_tkb = await sheet_tkb.getRows();
    // sort data
    API.count_api = 0;
    await funcSortData(stt_start ?? 0, sheet_data, data_vnedu, data_tkb, type);
    return handleSuccess(res, {}, "Thành công");
  } catch (err) {
    return handleError(res, "Lỗi không xác định", err);
  }
};

async function funcConvertData(
  index: number,
  data_tkb: any[],
  sheet_vnedu: any,
  map_data_gv: MAP_DATA_GV,
  map_data_mh: MAP_DATA_MH,
  teacher_current?: {
    ms_gv: string;
    name_gv: string;
  }
): Promise<void> {
  if (index >= data_tkb.length) return Promise.resolve();
  return new Promise(async (resolve, reject) => {
    const [name_gv, ca_hoc, mon, tong1, tong2, tong3] =
      data_tkb[index]._rawData;
    if (ca_hoc !== undefined && mon !== undefined) {
      const ms_gv = map_data_gv.get(name_gv)?.ms_gv ?? "";
      const name_gv_vnedu =
        map_data_gv.get(name_gv)?.gv_vnedu ??
        (map_data_gv.get(name_gv)?.ms_gv ? "" : "<NO USER>");
      API.count_api++;
      try {
        await sheet_vnedu.addRow({
          ["Mã số GV"]: ms_gv !== teacher_current?.ms_gv ? ms_gv : "",
          ["Họ tên / Tài khoản"]:
            name_gv_vnedu !== teacher_current?.name_gv ? name_gv_vnedu : "",
          ["Môn"]: mon,
          ["Các lớp dạy kỳ 1"]: tong1,
          ["Các lớp dạy kỳ 2"]: "",
        });
      } catch (err) {
        console.log("ERROR API");
        await new Promise((resolve) => setTimeout(resolve, 5000));
        console.log("RETRY API");
        await sheet_vnedu.addRow({
          ["Mã số GV"]: ms_gv !== teacher_current?.ms_gv ? ms_gv : "",
          ["Họ tên / Tài khoản"]:
            name_gv_vnedu !== teacher_current?.name_gv ? name_gv_vnedu : "",
          ["Môn"]: mon,
          ["Các lớp dạy kỳ 1"]: tong1,
          ["Các lớp dạy kỳ 2"]: "",
        });
      }
      if (name_gv !== teacher_current?.name_gv) {
        teacher_current = {
          ms_gv: map_data_gv.get(name_gv)?.ms_gv ?? "",
          name_gv: name_gv,
        };
      }
    }
    return funcConvertData(
      index + 1,
      data_tkb,
      sheet_vnedu,
      map_data_gv,
      map_data_mh,
      teacher_current
    ).then(() => {
      resolve();
    });
  });
}

export const createExcelVnedu = async (req, res) => {
  const VNEDU_TEMPLETE = "1x6gz9PpEdF8R6pSmzsOHscOfMj3j1iLj-mHNlyQCEjo";
  const { sheet_id_vnedu, sheet_id_tkb, sheet_id_data } = req.body;
  try {
    const vnedu_templete = new GoogleSpreadsheet(
      VNEDU_TEMPLETE,
      serviceAccountAuth
    );
    const doc_tkb = new GoogleSpreadsheet(sheet_id_tkb, serviceAccountAuth);
    const doc_vnedu = new GoogleSpreadsheet(sheet_id_vnedu, serviceAccountAuth);
    const doc_data = new GoogleSpreadsheet(sheet_id_data, serviceAccountAuth);
    await vnedu_templete.loadInfo();
    await doc_tkb.loadInfo();
    await doc_data.loadInfo();

    let sheet_vnedu_temp, sheet_data_gv, sheet_data_mh, sheet_vnedu, sheet_tkb;
    sheet_vnedu_temp = vnedu_templete.sheetsByIndex[0];
    for (let i = 0; i < doc_data.sheetCount; i++) {
      switch (doc_data.sheetsByIndex[i]?.title?.toUpperCase()) {
        case `DATA_GV`:
          sheet_data_gv = doc_data.sheetsByIndex[i];
          break;
        case `DATA_MH`:
          sheet_data_mh = doc_data.sheetsByIndex[i];
          break;
      }
    }

    await sheet_vnedu_temp.copyToSpreadsheet(doc_vnedu.spreadsheetId);
    await doc_vnedu.loadInfo();
    sheet_vnedu = doc_vnedu.sheetsByIndex[doc_vnedu.sheetCount - 1];
    sheet_vnedu.setHeaderRow(
      [
        "Mã số GV",
        "Họ tên / Tài khoản",
        "Môn",
        "Các lớp dạy kỳ 1",
        "Các lớp dạy kỳ 2",
      ],
      11
    );

    for (let i = 0; i < doc_tkb.sheetCount; i++) {
      await doc_tkb.sheetsByIndex[i].updateProperties({
        title: `tbk_${i}`,
      });
    }
    sheet_tkb = doc_tkb.sheetsByIndex[0];
    await sheet_tkb.setHeaderRow(
      ["Giáo viên", "Ca học	Môn", "Dạy cho lớp", "Tổng1", "Tổng2", "Tổng3"],
      5
    );

    const data_gv = await sheet_data_gv.getRows();
    const data_mh = await sheet_data_mh.getRows();
    const map_data_gv: MAP_DATA_GV = new Map();
    const map_data_mh: MAP_DATA_MH = new Map();
    data_gv.forEach((item) => {
      map_data_gv.set(item.get("GV_TKB"), {
        ms_gv: item.get("MA_SO_GV"),
        gv_vnedu: item.get("GV_VNEDU"),
        gv_tkb: item.get("GV_TKB"),
      });
    });
    data_mh.forEach((item) => {
      map_data_mh.set(item.get("MH_TKB"), {
        mh_vnedu: item.get("MH_VNEDU"),
        mh_tkb: item.get("MH_TKB"),
      });
    });
    const data_tkb = await sheet_tkb.getRows();
    await funcConvertData(0, data_tkb, sheet_vnedu, map_data_gv, map_data_mh);

    return handleSuccess(res, {}, "Thành công");
  } catch (err) {
    return handleError(res, "Lỗi không xác định", err);
  }
};
export const excelController = {
  createExcelVnedu,
  sortData,
};
