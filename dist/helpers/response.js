"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.getCurrentId = exports.errorSystem = exports.handleError = exports.handleSuccess = void 0;
const jsonwebtoken_1 = __importDefault(require("jsonwebtoken"));
const handleSuccess = (res, data, message) => {
    return res.send({
        code: 1,
        message: message,
        data: data,
    });
};
exports.handleSuccess = handleSuccess;
const handleError = (res, message, status) => {
    if (status) {
        res.status(status);
    }
    return res.send({
        code: 2,
        message: message,
    });
};
exports.handleError = handleError;
const errorSystem = (req, res, err) => {
    if (res)
        res.send({
            status: 0,
            code: 0,
            message: err.message,
        });
};
exports.errorSystem = errorSystem;
const getCurrentId = (req) => __awaiter(void 0, void 0, void 0, function* () {
    return new Promise((resolve, reject) => {
        var id = "";
        var token = req.body.token ||
            req.query.token ||
            req.headers["x-access-token"] ||
            req.cookies.token;
        if (token && token.search("token=") !== -1) {
            token = token.substring(token.search("token=") + 6);
        }
        if (!token) {
            id = "";
        }
        else {
            jsonwebtoken_1.default.verify(token, "minigames", function (err, decoded) {
                if (err) {
                    id = "";
                }
                else {
                    id = decoded.id;
                }
            });
        }
        resolve(id);
    });
});
exports.getCurrentId = getCurrentId;
//# sourceMappingURL=response.js.map