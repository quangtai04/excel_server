import express from "express";
import compression from "compression"; // compresses requests
import bodyParser from "body-parser";
import lusca from "lusca";
import session from "express-session";
import flash from "express-flash";
import cors from "cors";
import logger from "morgan";
import cookieParser from "cookie-parser";
import { router } from "./routers";
import { SESSION_SECRET } from "./util/secrets";

// Create Express server
const app = express();
const options: cors.CorsOptions = {
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
app.use(cors(options));

app.use(logger("dev"));
app.use(express.static("public"));

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));
// Express configuration
app.set("port", 3011);
logger.token("id", function getId(req: any) {
  return req.id;
});
app.use(compression());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cookieParser());
app.use(
  session({
    resave: true,
    saveUninitialized: true,
    secret: SESSION_SECRET,
  })
);
app.use(flash());
app.use(lusca.xframe("SAMEORIGIN"));
app.use(lusca.xssProtection(true));
app.use((req, res, next) => {
  next();
});
app.use((req, res, next) => {
  next();
});

router(app);

export default app;
