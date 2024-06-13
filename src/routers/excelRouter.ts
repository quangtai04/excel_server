import express from "express";
import * as excelController from "../controllers/excel";
const router = express.Router();

router.post("/excel2json", excelController.parserExcel2Json);

export default router;
