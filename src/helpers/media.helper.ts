import crypto from "crypto";
export const createHashMD5 = (text : string) => {
  var md5 = crypto.createHash("md5");
  md5.update(text);
  return md5.digest("hex");
};
