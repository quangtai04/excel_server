"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.domain = void 0;
const domain = () => {
    return process.env.NODE_ENV ? `localhost:${process.env.PORT}` : 'domain';
};
exports.domain = domain;
//# sourceMappingURL=domain.js.map