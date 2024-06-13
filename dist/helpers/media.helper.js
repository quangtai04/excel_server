"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.createHashMD5 = void 0;
const crypto_1 = __importDefault(require("crypto"));
const createHashMD5 = (text) => {
    var md5 = crypto_1.default.createHash("md5");
    md5.update(text);
    return md5.digest("hex");
};
exports.createHashMD5 = createHashMD5;
//# sourceMappingURL=media.helper.js.map