"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.handleStatus = void 0;
const handleStatus = (status, message, data) => {
    switch (status) {
        case 200:
            return {
                status,
                message: message || "Thành công",
                data,
            };
        case 302:
            return {
                status,
                message: message || "Đã tồn tại",
            };
        case 400:
            return {
                status,
                message: message || "Dữ liệu không hợp lệ",
            };
        case 404:
            return {
                status,
                message: message || "Không tìm thấy",
            };
        default:
            return { status: 500, message: message || "Thất bại" };
    }
};
exports.handleStatus = handleStatus;
//# sourceMappingURL=statusResponse.js.map