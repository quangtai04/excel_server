const { GoogleSpreadsheet } = require("google-spreadsheet");
const { JWT } = require("google-auth-library");
const { google } = require("googleapis");
const fs = require("fs");
const path = require("path");
const process = require("process");
import { xlsxParser } from "../util/xlsxParser";

import { handleError, handleSuccess } from "../helpers/response";

const GOOGLE_SERVICE_ACCOUNT_EMAIL =
  "download-excel@excelserver.iam.gserviceaccount.com";
const GOOGLE_PRIVATE_KEY =
  "-----BEGIN PRIVATE KEY-----\nMIIEugIBADANBgkqhkiG9w0BAQEFAASCBKQwggSgAgEAAoIBAQDKgqsPvYzO6FwB\ntp4N85nmXABGBm6VazAxJzJZK9/F+7ry7COwe18qGjuX/uBFk5r7mpAWQKVlRJnP\nKTnDPUwD99hN/nVG5K+Qv7TC9R916DJ/nxCD3aaKHHEFPvEHf9K4XcdVaXFKsB6s\n/WJurOJkP2rV3IzPDo0HqNobwx2bPyjXy6zt7CRIpMgNHBVLvzDudgXc/L/A8iXN\nhdZhiZ89FDqEDXcCua7zevj6bT2VAlLtihUNu4/Ui3gQ6m0/m6RLfnKmMxhcNt11\nHw6BgI8ZKuYOjK8Jm/VE/4KuJWOKX5TTpme+s3o1UUfEswXCbYID89AtOOSiYya9\nu5Yi0q4bAgMBAAECggEAMEFXwr3vriQXPH9IBVoNS6OTmxwQQMGUb7n/2NjID1TP\nNPCJBpY3VICAv9THmzyzew3XFL5dyxZAMmmH7pqOIQnvfJJMXtLCdRMBX01qrD1i\nvx9nn4xzEUj6s33OkHNogm8yPwuLp/j0rlMIoAfJQIsOCZzu3q4AdBlLs77YMLQ9\nH8IfBt+OY+dd2KPA58k1J2LdXo8MAf2Sqw3Ks5wOSZ7Wvewa+Gk/9n/X+cawUSR/\nI44yJnM5+33Udez3ol2Bq/IV39jQFZCW6S4Tka0iqoJfpH05WnTjEiy8nsCGs93y\nIlcOulExesJPkuufmjDb0z9YmNGg3KQ1Vn5Pp4e5mQKBgQDnkBYI68PhWCmpQh1b\nFjtEltx6+JHn3bOPTp8+ksP+Cl4Y6DEsQD2WdKY9+zn5XYrpHXLmuHdu2FL56zvn\nQJZgwdkUwwhPN7veuEBWGcsrHdAP0Gyl7nkU9oI028l30beevWH+5qe2+Gx0agYN\n0zfXkSuYe2CifZbHgTld+69S2QKBgQDf4bNZY1EVbuxqdFDa66OHnUcOsUwjZqpJ\nZ+slYwJJAFV9NUrje+r2iqdIEG0wd0mCMDJdgm4+TObIddYJnh6lzDBVHC26t/9a\nFThHaawfeTX3e0f2jfCqC4cNja/N2NMj5G37GoVA/zXy3aTFi65AeYAZdqVv9zmQ\nyfG/AwbIEwKBgHTnOxiZ3jQfzDiVFjjsClPgTcMPRqnmNUZ+DMsMUUIpfcPZRSnv\n7KfOkDbuZCBOZ1i081MjgbhGIe3mIkHnS4PhmXRv2fwUSRZxsplFQFquHGI/ePp0\nYFCC+s0wwI9rIuZS+ew0CivHUwmalR/ZqHF96qJ6dxjRipOB27Jk4+hRAoGAERue\nKOZ9+7VOO5RH2XLIPES4eVbzCoF94b4fKew28H0mCztXTaraeZx+y/L1ZQ64f0pP\njvW4luopeIgIoxOCJAlGaDqPBAWrbabRdiONE5qflRnGlgCis1vOJir9lC1NdSCv\nhtCv/heCd9yYCsBxwFuIfmmimru5mmbUIlSI4CECf2GLo/ZUiU/Om54yMxNwxlFa\nexQPw3owo/YEP1xbBdjIBXv9bI2N022orsgqa6/U487rFPG+MTL2yzIZoPuWuVsn\nplgMb9Dio40JnhQhe2jB7f3g+LFfYqecUXd15Rag34Qo4Q5aNQvU8RJEaPrfRyc7\nceEAbAOk/ugE4fz9Ggo=\n-----END PRIVATE KEY-----\n";

const serviceAccountAuth = new JWT({
  email: GOOGLE_SERVICE_ACCOUNT_EMAIL,
  key: GOOGLE_PRIVATE_KEY,
  scopes: [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
  ],
});

const getPath = (name) => {
  return path.join(process.cwd(), "/public/excel", name);
};

const download_file = async (
  fileId: string,
  fileName: string
): Promise<any> => {
  const drive = google.drive({ version: "v3", auth: serviceAccountAuth });
  if (!fs.existsSync(path.join(process.cwd(), "/public"))) {
    await fs.mkdirSync(path.join(process.cwd(), "/public"), {
      recursive: true,
    });
  }
  if (fs.existsSync(path.join(process.cwd(), "/public/excel"))) {
    await fs.mkdirSync(path.join(process.cwd(), "/public/excel"), {
      recursive: true,
    });
  }
  const file = fs.createWriteStream(getPath(fileName));
  return new Promise((resolve, reject) => {
    drive.files.get(
      { fileId: fileId, alt: "media" },
      { responseType: "stream" },
      (err, { data }) => {
        if (err) {
          reject(err);
          return;
        }
        data
          .on("end", () => {
            resolve(file);
          })
          .on("error", (err) => {
            reject(err);
            return process.exit();
          })
          .pipe(file);
      }
    );
  });
};

const upload_file = async (
  folderId: string,
  fileName: string,
  pathFile: string,
  isUnlink?: boolean
): Promise<string> => {
  const drive = google.drive({ version: "v3", auth: serviceAccountAuth });
  const fileMetadata = {
    name: fileName, // Tên tệp bạn muốn tải lên
    parents: [folderId], // Thêm ID của thư mục vào đây
  };
  const media = {
    mimeType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // MIME type của tệp xlsx
    body: fs.createReadStream(pathFile), // Đường dẫn đến tệp cục bộ
  };
  return new Promise((resolve, reject) => {
    drive.files.create(
      {
        resource: fileMetadata,
        media: media,
        fields: "id",
      },
      (err, file) => {
        if (err) {
          // Xử lý lỗi
          console.error(err);
        } else {
          resolve(file.data.id);
        }
        if (isUnlink) fs.unlinkSync(pathFile);
      }
    );
  });
};

export const renderFileData = async (req, res) => {
  const { folderId } = req.body;
  const drive = google.drive({ version: "v3", auth: serviceAccountAuth });
  const resDrive = await drive.files.list({
    q: `'${folderId}' in parents and trashed = false`,
  });
  const files = resDrive.data.files;
  let file_vnedu, file_tkb;
  files.forEach((file) => {
    switch (file.name) {
      case "vnedu_gv_mh.xlsx":
        file_vnedu = file;
        break;
      case "tkb.xlsx":
        file_tkb = file;
        break;
    }
  });
  const time = new Date().getTime().toString();
  const file_name_vnedu = `vnedu_gv_mh_${time}.xlsx`;
  const file_name_tkb = `tkb_${time}.xlsx`;
  await download_file(file_vnedu.id, file_name_vnedu);
  await download_file(file_tkb.id, file_name_tkb);
  await xlsxParser.renderData(
    getPath(file_name_vnedu),
    getPath(file_name_tkb),
    getPath(`data_${time}.xlsx`)
  );
  await fs.unlinkSync(getPath(file_name_vnedu));
  await fs.unlinkSync(getPath(file_name_tkb));
  await upload_file(folderId, `data.xlsx`, getPath(`data_${time}.xlsx`), true);
  return handleSuccess(res, { files }, "Thành công");
};

export const createExcelVnedu = async (req, res) => {
  const { folderId, isHK2 } = req.body;
  const drive = google.drive({ version: "v3", auth: serviceAccountAuth });
  const resDrive = await drive.files.list({
    q: `'${folderId}' in parents and trashed = false`,
  });
  const files = resDrive.data.files;
  let file_tkb, file_data, file_templete;
  files.forEach((file) => {
    switch (file.name) {
      case "tkb.xlsx":
        file_tkb = file;
        break;
      case "data.xlsx":
        file_data = file;
        break;
      case "templete.xlsx":
        file_templete = file;
        break;
    }
  });
  if (!file_tkb || !file_data || !file_templete) {
    return handleError(res, "File không tồn tại");
  }
  const time = new Date().getTime().toString();
  const file_name_tkb = `tkb_${time}.xlsx`;
  const file_name_data = `data_${time}.xlsx`;
  const file_name_result = `result_${time}.xlsx`;
  await download_file(file_tkb.id, file_name_tkb);
  await download_file(file_data.id, file_name_data);
  await download_file(file_templete.id, file_name_result);
  if (
    !fs.existsSync(getPath(file_name_result)) ||
    !fs.existsSync(getPath(file_name_data)) ||
    !fs.existsSync(getPath(file_name_tkb))
  ) {
    return handleError(res, "File không tồn tại");
  }
  await xlsxParser.convertTKBToVNEDU(
    getPath(file_name_tkb),
    getPath(file_name_data),
    getPath(file_name_result),
    isHK2
  );
  await fs.unlinkSync(getPath(file_name_tkb));
  await fs.unlinkSync(getPath(file_name_data));
  await upload_file(
    folderId,
    "vnedu_result.xlsx",
    getPath(file_name_result),
    true
  );
  return handleSuccess(res, {}, "Thành công");
};

export const excelController = {
  createExcelVnedu,
  renderFileData,
};
