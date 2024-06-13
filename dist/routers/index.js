"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.router = void 0;
const excelRouter_1 = __importDefault(require("./excelRouter"));
const router = (app) => {
    app.use("/api/excel", excelRouter_1.default);
};
exports.router = router;
//# sourceMappingURL=index.js.map