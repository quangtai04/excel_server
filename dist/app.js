"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const compression_1 = __importDefault(require("compression")); // compresses requests
const body_parser_1 = __importDefault(require("body-parser"));
const lusca_1 = __importDefault(require("lusca"));
const express_session_1 = __importDefault(require("express-session"));
const express_flash_1 = __importDefault(require("express-flash"));
const cors_1 = __importDefault(require("cors"));
const morgan_1 = __importDefault(require("morgan"));
const cookie_parser_1 = __importDefault(require("cookie-parser"));
const routers_1 = require("./routers");
const secrets_1 = require("./util/secrets");
// Create Express server
const app = (0, express_1.default)();
const options = {
    allowedHeaders: [
        "Origin",
        "X-Requested-With",
        "Content-Type",
        "Accept",
        "X-Access-Token",
    ],
    credentials: true,
    methods: "GET,HEAD,OPTIONS,PUT,PATCH,POST,DELETE",
    origin: "*",
    preflightContinue: false,
};
app.use((0, cors_1.default)(options));
app.use((0, morgan_1.default)("dev"));
app.use(express_1.default.static("public"));
app.use(body_parser_1.default.json());
app.use(body_parser_1.default.urlencoded({ extended: false }));
// Express configuration
app.set("port", 3011);
morgan_1.default.token("id", function getId(req) {
    return req.id;
});
app.use((0, compression_1.default)());
app.use(body_parser_1.default.json());
app.use(body_parser_1.default.urlencoded({ extended: true }));
app.use((0, cookie_parser_1.default)());
app.use((0, express_session_1.default)({
    resave: true,
    saveUninitialized: true,
    secret: secrets_1.SESSION_SECRET,
}));
app.use((0, express_flash_1.default)());
app.use(lusca_1.default.xframe("SAMEORIGIN"));
app.use(lusca_1.default.xssProtection(true));
app.use((req, res, next) => {
    next();
});
app.use((req, res, next) => {
    next();
});
(0, routers_1.router)(app);
exports.default = app;
//# sourceMappingURL=app.js.map